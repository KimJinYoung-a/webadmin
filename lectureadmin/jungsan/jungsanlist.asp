<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
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
dim tot1,tot2,tot3,totsum
tot1 = 0
tot2 = 0
tot3 = 0
totsum = 0
%>
<script language='javascript'>
function PopDetail(mnupos,iidx){
	var popwin = window.open('popshowdetail.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=800, height=540, scrollbars=1');
	popwin.focus();
}

function PopConfirm(mnupos,iidx){
	var popwin = window.open('jungsanmaster.asp?id=' + iidx + '&menupos=' + mnupos,'popshowdetail','width=800, height=540, scrollbars=1');
	popwin.focus();
}

function PopTaxReg(v){
	var popwin = window.open("/designer/jungsan/poptaxreg.asp?id=" + v,"poptaxreg","width=640 height=680 scrollbars=yes resizable=yes");
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
</script>



<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr valign="bottom">
    <td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
    <td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
  </tr>
  <tr valign="top">
    <td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom">
    <font color="red"><strong>공지사항[필독]</strong></font></td>
    <td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td align="right" bgcolor="#F3F3FF">
	        <a href="http://webadmin.10x10.co.kr/images/10x10lic.jpg" target="_blank"><font color="#0000FF">[사업자등록증보기]</font></a>
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF">
		<strong>전자세금계산서</strong> 발행을 원칙으로 합니다.
		<br>정산내역을 확인하시고, 오른쪽 끝에 있는 <strong>'전자세금계산서발행'</strong>을 클릭하시고, 안내설명에 따라 발행하시면 됩니다.
		<br>빠른 내역처리를 위해 꼭 텐바이텐어드민에서 발행키를 눌러 발행하시기 바랍니다.(사업자인증 불필요)
		<br>(네오포트사이트에서 발행하시면 안됩니다.)
		<br>발행하시기 전에 꼭 어드민에 등록되어 있는 <strong>업체정보를 확인</strong>하시고, 내용이 상이할 경우 수정바랍니다.(정산담당자 필수입력)
		<br>담당자 : 이문재 팀장 &nbsp;&nbsp;
		    문의전화 : 1644-1851(806) / 554-2033(104)&nbsp;017-360-6991&nbsp;&nbsp;
		    문의메일 : <a href="mailto:moon@10x10.co.kr">moon@10x10.co.kr</a>
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>

  <tr valign="top" bgcolor="#F3F3FF">
    <td height="10" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td height="10" background="/images/tbl_blue_round_08.gif"></td>
    <td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
  </tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#bababa">
    <tr align="center" bgcolor="#DDDDFF">
      	<td width="90">Title</td>
      	<td width="24">차수</td>
      	<td width="24">과세</td>
      	<td width="64">업체배송<br>총액</td>
      	<td width="64">매입총액</td>
      	<td width="64">특정총액</td>
      	<td width="64">총정산액</td>
<!--
      	<td width="65">등록일</td>
-->
      	<td width="70">발행일</td>
      	<td width="70">입금일</td>
      	<td width="80">상태</td>
      	<td width="50">상세내역</td>
      	<td>전자계산서<br>발행</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	tot1 = tot1 + ojungsan.FItemList(i).Fub_totalsuplycash
    	tot2 = tot2 + ojungsan.FItemList(i).Fme_totalsuplycash
    	tot3 = tot3 + ojungsan.FItemList(i).Fwi_totalsuplycash + ojungsan.FItemList(i).Fet_totalsuplycash + ojungsan.FItemList(i).Fsh_totalsuplycash
    %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td>
      		<a href="javascript:PopConfirm('<%= menupos %>','<%= ojungsan.FItemList(i).FId %>')"><%= ojungsan.FItemList(i).Ftitle %>
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
     	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).Fwi_totalsuplycash + ojungsan.FItemList(i).Fet_totalsuplycash + ojungsan.FItemList(i).Fsh_totalsuplycash,0) %></td>
      	<td align="right"><%= FormatNumber(ojungsan.FItemList(i).GetTotalSuplycash,0) %></td>
<!--
      	<td><%= Left(Cstr(ojungsan.FItemList(i).FRegDate),10) %></td>
-->
      	<td><%= ojungsan.FItemList(i).Ftaxregdate %></td>
      	<td><%= ojungsan.FItemList(i).Fipkumdate %></td>
      	<td><font color="<%= ojungsan.FItemList(i).GetStateColor %>"><%= ojungsan.FItemList(i).GetStateName %></font></td>
      	<td>
      		<a href="javascript:PopDetail('<%= menupos %>','<%= ojungsan.FItemList(i).FId %>');">보기
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
    <% totsum = totsum + tot1 + tot2 + tot3 %>
    <% if ojungsan.FResultCount<1 then %>
    <tr bgcolor="#FFFFFF">
      	<td align="center" colspan="12">검색결과가 없습니다.</td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
      	<td >합계</td>
      	<td ></td>
      	<td ></td>
      	<td align="right"><%= FormatNumber(tot1,0) %></td>
      	<td align="right"><%= FormatNumber(tot2,0) %></td>
      	<td align="right"><%= FormatNumber(tot3,0) %></td>
      	<td align="right"><%= FormatNumber(totsum,0) %></td>
<!--
      	<td ></td>
-->
      	<td ></td>
      	<td ></td>
      	<td ></td>
      	<td ></td>
      	<td ></td>
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