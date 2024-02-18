<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/offjungsancls.asp"-->

<%

dim makerid, finishflag, page
makerid = session("ssBctId")
finishflag = request("finishflag")
page = request("page")

if page="" then page=1

dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FPageSize   = 100
ooffjungsan.FCurrPage = page
ooffjungsan.FRectMakerid = makerid
ooffjungsan.GetOffJungsanMasterListBrandView


dim i
dim orgsellmargin, realsellmargin
orgsellmargin   = 0
realsellmargin  = 0
%>
<script language='javascript'>
function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function PopDetail(idx){
    var popwin = window.open('off_jungsandetailsum.asp?idx=' + idx ,'off_jungsandetailsum','width=800, height=540, scrollbars=yes, resizable=yes');
    popwin.focus();
}

function PopTaxRegOff(v){
	var popwin = window.open("/designer/jungsan/poptaxregoff.asp?idx=" + v,"poptaxregoff1","width=640 height=680 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopTaxPrint(itax_no,ibizno){
	var popwinsub = window.open("http://www.neoport.net/jsp/dti/tx/dti_get_pin.jsp?tax_no=" + itax_no + "&cur_biz_no=" + ibizno,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

function PopTaxPrintReDirect(itax_no){
	var popwinsub = window.open("/designer/jungsan/red_taxprint.asp?tax_no=" + itax_no ,"taxview","width=670,height=620,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwinsub.focus();
}

alert('2014년 1월 정산내역은 수수료 정산 변경 작업 관계로\n\n2월 4일 오픈 되오니 양해 부탁드립니다.');
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top" bgcolor="F4F4F4" width="760"></td>
        <td valign="top" bgcolor="F4F4F4" >&nbsp;</td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="<%= adminColor("topbar") %>">
        <td colspan="20" align="right" >
            총건수: <%= FormatNumber(ooffjungsan.FTotalCount,0) %> &nbsp;&nbsp;
            총금액: <%= FormatNumber(ooffjungsan.FTotalSum,0) %> &nbsp;&nbsp;
            Page: <%= page %>/<%= ooffjungsan.FTotalPage %>
        </td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td>Title</td>
      <td width="25">차수</td>
      <td width="25">과세</td>
      <td width="50">특정<br>판매</td>
      <td width="50">업체<br>특정</td>
      <td width="50">오프<br>매입</td>
      <td width="50">매장<br>매입</td>
      <td width="50">출고<br>매입</td>
      <td width="50">기타<br>내역</td>
      <td width="66">총정산액</td>
      <td width="65">세금계산서<br>등록일</td>
      <td width="65">세금계산서<br>발행일</td>
      <td width="65">입금일</td>
      <td width="70">상태</td>
      <td width="50">상세<br>내역</td>
      <td width="80">전자계산서<br>발행</td>
    </tr>
    <% if ooffjungsan.FResultCount>0 then %>
    <% for i=0 to ooffjungsan.FResultCount-1 %>
    <%
        if (ooffjungsan.FItemList(i).Ftot_orgsellprice<>0) then
            orgsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_orgsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_orgsellprice*100*100)/100
        else
            orgsellmargin = 0
        end if

        if (ooffjungsan.FItemList(i).Ftot_realsellprice<>0) then
            realsellmargin = CLng((ooffjungsan.FItemList(i).Ftot_realsellprice-ooffjungsan.FItemList(i).Ftot_jungsanprice)/ooffjungsan.FItemList(i).Ftot_realsellprice*100*100)/100
        else
            realsellmargin = 0
        end if
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td align="left"><a href="javascript:PopDetail('<%= ooffjungsan.FItemList(i).Fidx %>');"><%= ooffjungsan.FItemList(i).Ftitle %></a></td>
      	<td ><%= ooffjungsan.FItemList(i).Fdifferencekey %></td>
      	<td ><font color="<%= ooffjungsan.FItemList(i).GetTaxtypeNameColor %>"><%= ooffjungsan.FItemList(i).GetSimpleTaxtypeName %></font></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FTW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FUW_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FOM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FSM_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FCM_price,0) %></td>
        <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FET_price,0) %></td>
      	<td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></td>
      	<td >
      	    <% if IsNULL(ooffjungsan.FItemList(i).Ftaxinputdate) then %>
			&nbsp;
	   	  	<% else %>
	     	<%= Left(Cstr(ooffjungsan.FItemList(i).Ftaxinputdate),10) %>
	      	<% end if %>
      	</td>
      	<td ><%= ooffjungsan.FItemList(i).Ftaxregdate %></td>
      	<td ><%= ooffjungsan.FItemList(i).Fipkumdate %></td>
      	<td ><font color="<%= ooffjungsan.FItemList(i).GetStateColor %>"><%= ooffjungsan.FItemList(i).GetStateName %></font></td>
      	<td ><a href="javascript:PopDetail('<%= ooffjungsan.FItemList(i).Fidx %>');">보기<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom"></a></td>
      	<td>
      	    <% if ooffjungsan.FItemList(i).IsElecTaxExists then %>
          	<a href="javascript:PopTaxPrintReDirect('<%= ooffjungsan.FItemList(i).Fneotaxno %>');">계산서출력
          	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
          	</a>
          	<% elseif ooffjungsan.FItemList(i).IsElecTaxCase then %>
          	<a href="javascript:PopTaxRegOff('<%= ooffjungsan.FItemList(i).FIdx %>');">계산서발행
          	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
          	</a>
          	<% elseif ooffjungsan.FItemList(i).IsElecFreeTaxCase then %>
          	<a href="javascript:PopTaxRegOff('<%= ooffjungsan.FItemList(i).FIdx %>');">계산서발행
          	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
          	</a>
          	<% end if %>
      	</td>
    </tr>
    <% next %>
    <% else %>
    <tr bgcolor="#FFFFFF">
		<td colspan=20 align="center">[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
			<% if ooffjungsan.HasPreScroll then %>
				<a href="javascript:NextPage('<%= ooffjungsan.StartScrollPage-1 %>')">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ooffjungsan.StartScrollPage to ooffjungsan.FScrollCount + ooffjungsan.StartScrollPage - 1 %>
				<% if i>ooffjungsan.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ooffjungsan.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set ooffjungsan = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->