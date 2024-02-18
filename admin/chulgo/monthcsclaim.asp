<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간 CS문의 및 클레임(~주)
' History : 2007.08.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->


<%
response.write "사용중지"
response.end

dim yyyy,graphyyyy
	yyyy = request("yyyy")
		if (yyyy="") then yyyy = Cstr(Year(now()))
graphyyyy = request("graphyyyy")


session("yyyy") = yyyy

dim omonthcsclaim , i
	set omonthcsclaim = new Cchulgoitemlist
	omonthcsclaim.frectyyyy = yyyy
	omonthcsclaim.fmonthcsclaim()

dim omonthcssangdam 
	set omonthcssangdam = new Cchulgoitemlist
	omonthcssangdam.frectyyyy = yyyy
	omonthcssangdam.fmonthcssangdam()	
%>

<script language="javascript">

<!--cs유형별클레임통계를위한 팝업링크 시작-->
function jsyyyy(yyyy)
{
var popup = window.open('/admin/chulgo/monthcsclaim_detail.asp?yyyy='+yyyy,'jsyyyy','width=1024,height=768,scrollbars=yes,resizable=yes');
popup.focus();
}
<!--cs유형별클레임통계를위한 팝업링크 끝-->

function submit()
{
document.frm.submit();
}
function chk(){
	document.frm.graphyyyy.value = document.graph.graphyyyy.value;
	document.frm.submit();
}
</script>
<script language="javascript" src="/admin/chulgo/daumchart/FusionCharts.js"></script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>월간CS문의 및 클레임</strong></font>
			</td>
			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>

</table>
<!--표 헤드끝-->

<!-- 표 검색부분 시작-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	
	<form name="frm" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	
		<tr bgcolor="#FFFFFF" valign="top">
	        <td background="/images/tbl_blue_round_04.gif" width="1%" bgcolor="F4F4F4"></td>
	        <td width="54%" bgcolor="F4F4F4"> 
	       		년: &nbsp;<% DrawYBox yyyy %>
	        	<input type="submit" value="검색">
	        </td>
	        <td valign="top" align="right" width="40%" bgcolor="F4F4F4">
	      	</td>
	        <td background="/images/tbl_blue_round_05.gif" bgcolor="F4F4F4" width="1%"></td>
	    </tr>
    </form>
    
	<!--<form name="graph" method="get" action="/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).asp">
	<tr><td><input type="text" name="graphyyyy" value="<%=yyyy%>"></td></tr>
	</form>-->
</table>
<!-- 표 검색부분 끝-->		

<!-- 월간 접수건수 시작-->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	
<% if omonthcsclaim.ftotalcount > 0 then %>
	<tr>
		<td bgcolor="ffffff" colspan=9>
		월간 접수건수
		</td>
	</tr>
	<tr bgcolor=#DDDDFF>
		<td align="center">달 | 유형</td>
		<td align="center">맞교환출고</td>
		<td align="center">누락재발송</td>
		<td align="center">서비스발송</td>
		<td align="center">반품</td>
		<td align="center">회수</td>
		<td align="center">맞교환회수</td>
		<td align="center">주문취소</td>
		<td align="center">합계</td>
	</tr>
	<% dim fitemtotal %>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><a href="javascript:jsyyyy('<%= omonthcsclaim.flist(i).fyyyy %>')"><%= omonthcsclaim.flist(i).fyyyy %></a></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd0 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd1 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd2 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd3 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd4 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd5 %></td>
			<td align="center"><%= omonthcsclaim.flist(i).fitemd6 %></td>
			<td align="center"><% fitemtotal = omonthcsclaim.flist(i).fitemd0+omonthcsclaim.flist(i).fitemd1+omonthcsclaim.flist(i).fitemd2+omonthcsclaim.flist(i).fitemd3+omonthcsclaim.flist(i).fitemd4+omonthcsclaim.flist(i).fitemd5+omonthcsclaim.flist(i).fitemd6 %>
			<%= fitemtotal %></td>
		</tr>
	<% next %>
<!-- 월간 접수건수 끝-->	
	</table>
	
	<!--그래프 통계 시작-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<!--<td align="center">
			<embed
			src="/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).swf?page_name=/admin/chulgo/SmartChart_line2/SmartChart_line2(Beta).asp&data_1=&data_2="
			quality="high" scale="noscale"
			bgcolor="#ffffff" width="800" height="600" name="barchart" align="middle" 
			allowScriptAccess="sameDomain" type="application/x-shockwave-flash"
			pluginspage="http://www.macromedia.com/go/getflashplayer">
			</embed> 
		</td>-->
		<td align="center">
		<div align="right"><input type="button" value="그래프프린트" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv3" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D.asp");
				chart.render("chartdiv3");
			</script>
		</td>
	</tr>
	</table><br>
	<!--그래프 통계 끝-->
	
<% else %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF"><%= yyyy %>년 월간접수 검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>
<% if omonthcssangdam.ftotalcount > 0 then %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td bgcolor="ffffff" colspan=11>
			1:1 상담
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">유형</td>
		<td align="center">배송</td>
		<td align="center">주문</td>
		<td align="center">상품</td>
		<td align="center">재고</td>
		<td align="center">취소</td>
		<td align="center">환불</td>
		<td align="center">교환</td>
		<td align="center">AS</td>
		<td align="center">이벤트</td>
		<td align="center">증빙서류</td>
		</tr>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd0 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd1 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd2 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd3 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd4 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd5 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd6 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd7 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd8 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd9 %></td>
		</tr>
	<% next %>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">유형</td>
		<td align="center">시스템</td>
		<td align="center">회원제도</td>
		<td align="center">회원정보</td>
		<td align="center">당첨</td>
		<td align="center">반품</td>
		<td align="center">입금</td>
		<td align="center">오프라인</td>
		<td align="center">쿠폰/마일리지</td>
		<td align="center">결제방법</td>
		<td align="center">기타</td>
		</tr>
	<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
		<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd10 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd11 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd12 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd13 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd14 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd15 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd16 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd17 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd18 %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemd20 %></td>
		</tr>
	<% next %>	
	</table>
	<!--사용안함 <table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td bgcolor="ffffff" colspan=11>
			1:1 상담합계 및 주문대상담비
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
		<td align="center">달</td>
		<td align="center">문의합계</td>
		<td align="center">주문대비</td>
		</tr>
		<% for i=0 to omonthcsclaim.FTotalCount - 1 %> 
			<tr bgcolor=#FFFFFF>
			<td align="center"><%= omonthcssangdam.flist(i).fyyyy %></td>
			<td align="center"><%= omonthcssangdam.flist(i).fitemdtot %></td>
			<td align="center">주문대비</td>
			</tr>
		<% next %>	
	</table>-->
	
	<!--1:1 상담 그래프 통계1 시작-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<td align="center">
		<div align="right"><input type="button" value="그래프프린트" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv4" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D1.asp");
				chart.render("chartdiv4");
			</script>
		</td>
	</tr>
	</table>
	<!--1:1 상담 그래프 통계1 끝-->
	<!--1:1 상담 그래프 통계2 시작-->
	<br>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr bgcolor=#FFFFFF>
		<td align="center">
		<div align="right"><input type="button" value="그래프프린트" onclick="javascript:window.print();"></div><br>
			<div id="chartdiv5" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/admin/chulgo/daumchart/MSCombiDY2D.swf", "chartdiv3", "800", "600", "0", "0");
				chart.setDataURL("/admin/chulgo/daumchart/MSCombiDY2D2.asp");
				chart.render("chartdiv5");
			</script>
		</td>
	</tr>
	</table>
	<!--1:1 상담 그래프 통계2 끝-->	
	<% else %>
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr align="center" bgcolor="#DDDDFF">
	<td align=center bgcolor="#FFFFFF"><%= yyyy %>년 1:1 상담 검색 결과가 없습니다.</td>
	</tr>
	</table>		
<% end if %>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->