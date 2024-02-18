<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim designer
designer = session("ssBctID")

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectDesigner = designer
ojungsan.FRectDesignerViewOnly = true
ojungsan.JungsanMasterList

dim i
dim tot1,tot2,tot3,tot4,totsum
tot1 = 0
tot2 = 0
tot3 = 0
tot4 = 0
totsum = 0
%>
<script language='javascript'>
function PopDetail(iidx){
	var popwin = window.open('jungsandetailsum.asp?id=' + iidx,'PopDetail','width=900, height=540, scrollbars=1');
	popwin.focus();
}

<!-- 사용안함 -->
function PopConfirm(mnupos,iidx){
	var popwin = window.open('jungsanmaster.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=900, height=540, scrollbars=1');
	popwin.focus();
}

function PopTaxReg(v){
	var popwin = window.open("poptaxreg.asp?id=" + v,"poptaxreg","width=640 height=700 scrollbars=yes resizable=yes");
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
    <tr height="10" valign="bottom" bgcolor="F4F4F4">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom" bgcolor="F4F4F4">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top" bgcolor="F4F4F4">&nbsp;</td>
        <td valign="top" bgcolor="F4F4F4"></td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      	<td width="120">Title</td>
      	<td width="24">차수</td>
      	<td width="24">과세</td>
      	<td width="64">업체배송<br>총액</td>
      	<td width="64">매입총액</td>
      	<td width="64">특정총액</td>
      	<td width="64">기타정산<br>총액</td>
      	<td width="64">총정산액</td>
      	<td width="70">세금계산서<br>등록일</td>
      	<td width="70">세금계산서<br>발행일</td>
      	<td width="70">입금일</td>
      	<td width="80">상태</td>
      	<td width="50">상세내역</td>
      	<td>전자계산서발행</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	tot1 = tot1 + ojungsan.FItemList(i).Fub_totalsuplycash
    	tot2 = tot2 + ojungsan.FItemList(i).Fme_totalsuplycash
    	tot3 = tot3 + ojungsan.FItemList(i).Fwi_totalsuplycash
    	tot4 = tot4 + ojungsan.FItemList(i).Fet_totalsuplycash + ojungsan.FItemList(i).Fsh_totalsuplycash
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td>
      		<a href="javascript:PopDetail('<%= ojungsan.FItemList(i).FId %>');"><%= ojungsan.FItemList(i).Ftitle %>
      		<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
        	</a>
      	</td>
      	<td><%= ojungsan.FItemList(i).Fdifferencekey %></td>
      	<td>
      		<% if ojungsan.FItemList(i).Ftaxtype="02" then %>
      		<font color=red>면세<font>
      		<% end if %>
      		<% if ojungsan.FItemList(i).Ftaxtype="01" then %>
      		과세
      		<% end if %>
      	</td>
     	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fub_totalsuplycash,0) %></td>
      	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fme_totalsuplycash,0) %></td>
     	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash,0) %></td>
      	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fet_totalsuplycash + ojungsan.FItemList(i).Fsh_totalsuplycash,0) %></td>
      	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
      	<td>
      		<% if IsNULL(ojungsan.FItemList(i).Ftaxinputdate) then %>
			&nbsp;
	   	  	<% else %>
	     	<%= Left(Cstr(ojungsan.FItemList(i).Ftaxinputdate),10) %>
	      	<% end if %>
      	</td>
      	<td><%= ojungsan.FItemList(i).Ftaxregdate %></td>
      	<td><%= ojungsan.FItemList(i).Fipkumdate %></td>
      	<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
      	<td>
      		<a href="javascript:PopDetail('<%= ojungsan.FItemList(i).FId %>');">보기
      		<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      		</a>
      	</td>
      	<td>
      	<% if ojungsan.FItemList(i).IsElecTaxExists then %>
      	<a href="javascript:PopTaxPrintReDirect('<%= ojungsan.FItemList(i).Fneotaxno %>');">(세금)계산서출력
      	<img src="/images/icon_print02.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsan.FItemList(i).IsElecTaxCase then %>
      	<a href="javascript:PopTaxReg('<%= ojungsan.FItemList(i).FId %>');">세금계산서발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsan.FItemList(i).IsElecFreeTaxCase then %>
      	<a href="javascript:PopTaxReg('<%= ojungsan.FItemList(i).FId %>');">계산서발행
      	<img src="/images/icon_new.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% elseif ojungsan.FItemList(i).IsElecSimpleBillCase then %>
      	<a href="javascript:PopConfirm('<%= menupos %>','<%= ojungsan.FItemList(i).FId %>');">정산확인
      	<img src="/images/icon_arrow_link.gif" width="14" height="14" border="0" align="absbottom">
      	</a>
      	<% end if %>
      	</td>
    </tr>
    <% next %>
    <% totsum = totsum + tot1 + tot2 + tot3 + tot4 %>
    <% if ojungsan.FResultCount<1 then %>
    <tr bgcolor="#FFFFFF">
      	<td align="center" colspan="15">검색결과가 없습니다.</td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td>합계</td>
      	<td></td>
      	<td></td>
      	<td align="right"><%= FormatNumber(tot1,0) %></td>
      	<td align="right"><%= FormatNumber(tot2,0) %></td>
      	<td align="right"><%= FormatNumber(tot3,0) %></td>
      	<td align="right"><%= FormatNumber(tot4,0) %></td>
      	<td align="right"><%= FormatNumber(totsum,0) %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
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
set ojungsan = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->