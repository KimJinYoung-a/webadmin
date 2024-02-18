<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanDiffcls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sellsite,yyyy1, mm1
dim nowDate, preDate, nextDate
dim clsJS, arrList, intLoop,arrActMeachul
	dim scm1, scm2, scm3, scm4, om1,om2,om3,om4
	dim scm1_d, scm2_d, scm3_d, scm4_d, om1_d,om2_d,om3_d,om4_d
	dim extscm, extom , extscm_d, extom_d
	dim extm, extc, extj
	dim itemmeachul(3), deliverymeachul(3)

   sellsite = requestCheckVar(Request("sellsite"),10)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
  if sellsite ="" then sellsite ="ssg"
   yyyy1 = "2018"
   mm1 = "03"
   if yyyy1="" then
   	nowDate = left(dateadd("m",-1,date()),7)
   	yyyy1 = left(nowDate,4)
   	mm1=mid(nowDate,6,2)
  else
  	nowDate =yyyy1&"-"&Format00(2,mm1)
	end if

   preDate = left(dateadd("m",-1,nowDate&"-01"),7)
   nextDate = left(dateadd("m",1,nowDate&"-01"),7)

   set clsJS = new CextJungsanMapping
   clsJS.FRectOutMall = sellsite
   clsJS.FRectyyyymm =nowDate
   arrList = clsJS.fnGetextMatchingData
   set clsJS = nothing


%>
<script type="text/javascript">
	function jsGoDetail(sellsite,stype, yyyy1,mm1,yyyy2,mm2){
		var popItem = window.open("extJungsanMappingItem.asp?sellsite="+sellsite+"&yyyy1="+yyyy1+"&mm1="+mm1+"&yyyy2="+yyyy2+"&mm2="+mm2+"&stype="+stype,"winItem","");
	}

	function jsExtJungsanDiffMake(){
		if(document.frmDB.sellsite.value ==""){
			alert("제휴몰을 선택해주세요");
			return;
		}

		document.frmDB.target="ifrDB";
		document.frmDB.submit();
	}
</script>
<!-- 검색 시작 -->

<form name="frmDB" method="post" action="extjungsanMappingProc.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sellsite" value="<%= sellsite %>">
<input type="hidden" name="jsdate" value="<%=nowDate%>">
</form>
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		제휴몰:	<% fnGetOptOutMall sellsite %>
		&nbsp;
		매출월:
		<% DrawYMBox yyyy1,mm1 %>

	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;

	</td>
</tr>
</table>
</form>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type="button" class="button" value="매출매핑(<%= sellsite %>, <%= nowDate %>)" onClick="jsExtJungsanDiffMake('<%= sellsite %>', '<%= nowDate %>');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>
<div width="2000" style="overflow-x:auto;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 <Tr bgcolor="#E6E6E6" align="center">
	<TD rowspan="2">제휴몰</TD>
	 <TD rowspan="2" colspan="2">구분</TD>
	 <TD colspan="3"><%=preDate%></TD>
	 <TD colspan="3"><%=nowDate%></TD>
	 <TD colspan="3"><%=nextDate%></TD>

	 <TD colspan="5">제휴사검토</TD>
	</tr>
	<Tr bgcolor="#E6E6E6" align="center">
		<TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>
	 	<td>미출고(10x10)</td>
	 	<td>미정산(제휴)</td>
	</tr>
	<% dim sumItemA(3), sumItemS(3), sumItemO(3)
	   dim sumDeA(3), sumDeS(3), sumDeO(3)
	   dim sumTotA(3), sumTotS(3), sumTotO(3)
	if isArray(arrList) then
		for intLoop = 0 To ubound(arrList,2)
		sumItemA(0) = sumItemA(0)+ arrList(1,intLoop)
		sumItemA(1) = sumItemA(1)+ arrList(2,intLoop)
		sumItemA(2) = sumItemA(2)+ arrList(3,intLoop)

		sumItemS(0) = sumItemS(0)+ arrList(4,intLoop)
		sumItemS(1) = sumItemS(1)+ arrList(5,intLoop)
		sumItemS(2) = sumItemS(2)+ arrList(6,intLoop)

		sumItemO(0) = sumItemO(0)+ arrList(7,intLoop)
		sumItemO(1) = sumItemO(1)+ arrList(8,intLoop)
		sumItemO(2) = sumItemO(2)+ arrList(9,intLoop)

		sumDeA(0) = sumDeA(0)+ arrList(10,intLoop)
		sumDeA(1) = sumDeA(1)+ arrList(11,intLoop)
		sumDeA(2) = sumDeA(2)+ arrList(12,intLoop)

		sumDeS(0) = sumDeS(0)+ arrList(13,intLoop)
		sumDeS(1) = sumDeS(1)+ arrList(14,intLoop)
		sumDeS(2) = sumDeS(2)+ arrList(15,intLoop)

		sumDeO(0) = sumDeO(0)+ arrList(16,intLoop)
		sumDeO(1) = sumDeO(1)+ arrList(17,intLoop)
		sumDeO(2) = sumDeO(2)+ arrList(18,intLoop)


		%>
	<tr bgcolor="#FFFFFF" align="right">
		<TD  rowspan="3" align="center"><%=sellsite%></TD>
		<td rowspan="3" align="center"><%=arrList(0,intLoop)%></td>
		<TD  bgcolor="#E6E6E6" width="50" align="center">상품</TD>
		<td><%if not isNull(arrList(1,intLoop)) then %><%=formatnumber(arrList(1,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(4,intLoop)) then %><%=formatnumber(arrList(4,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(7,intLoop)) then %><%=formatnumber(arrList(7,intLoop),0)%><%end if%></td>

		<td><%if not isNull(arrList(2,intLoop)) then %><%=formatnumber(arrList(2,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(5,intLoop)) then %><%=formatnumber(arrList(5,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(8,intLoop)) then %><%=formatnumber(arrList(8,intLoop),0)%><%end if%></td>

		<td><%if not isNull(arrList(3,intLoop)) then %><%=formatnumber(arrList(3,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(6,intLoop)) then %><%=formatnumber(arrList(6,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(9,intLoop)) then %><%=formatnumber(arrList(9,intLoop),0)%><%end if%></td>

		<td><%=formatnumber(arrList(3,intLoop),0)%></td>
		<td><%=formatnumber(arrList(6,intLoop),0)%></td>
		<td><%=formatnumber(arrList(9,intLoop),0)%></td>

		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(1,intLoop)+arrList(2,intLoop)+arrList(3,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(4,intLoop)+arrList(5,intLoop)+arrList(6,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(7,intLoop)+arrList(8,intLoop)+arrList(9,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((arrList(1,intLoop)+arrList(2,intLoop)+arrList(3,intLoop))-(arrList(4,intLoop)+arrList(5,intLoop)+arrList(6,intLoop)),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((arrList(1,intLoop)+arrList(2,intLoop)+arrList(3,intLoop))-(arrList(7,intLoop)+arrList(8,intLoop)+arrList(9,intLoop)),0)%></td>

	</tr>
	<Tr bgcolor="#FFFFFF" align="right">
	 	<TD  bgcolor="#E6E6E6" align="center">배송비</TD>
		<td><%if not isNull(arrList(10,intLoop)) then %><%=formatnumber(arrList(10,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(13,intLoop)) then %><%=formatnumber(arrList(13,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(16,intLoop)) then %><%=formatnumber(arrList(16,intLoop),0)%><%end if%></td>

		<td><%if not isNull(arrList(11,intLoop)) then %><%=formatnumber(arrList(11,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(14,intLoop)) then %><%=formatnumber(arrList(14,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(17,intLoop)) then %><%=formatnumber(arrList(17,intLoop),0)%><%end if%></td>

		<td><%if not isNull(arrList(12,intLoop)) then %><%=formatnumber(arrList(12,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(15,intLoop)) then %><%=formatnumber(arrList(15,intLoop),0)%><%end if%></td>
		<td><%if not isNull(arrList(18,intLoop)) then %><%=formatnumber(arrList(18,intLoop),0)%><%end if%></td>

		<td><%=formatnumber(arrList(12,intLoop),0)%></td>
		<td><%=formatnumber(arrList(15,intLoop),0)%></td>
		<td><%=formatnumber(arrList(18,intLoop),0)%></td>

	  	<td bgcolor="#EEEEEE"><%=formatnumber(arrList(10,intLoop)+arrList(11,intLoop)+arrList(12,intLoop),0)%> </td>
	  	<td bgcolor="#EEEEEE"><%=formatnumber(arrList(13,intLoop)+arrList(14,intLoop)+arrList(15,intLoop),0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(16,intLoop)+arrList(17,intLoop)+arrList(18,intLoop),0)%> </td>
		<td bgcolor="#EEEEEE"> </td>
		<td bgcolor="#EEEEEE"> </td>
	</tr>
	<Tr bgcolor="#FFFFFF" align="right">
	 	<TD  bgcolor="#E6E6E6" align="center">합계</TD>
		<td><%=formatnumber(arrList(1,intLoop)+arrList(10,intLoop),0)%></td>
		<td><%=formatnumber(arrList(4,intLoop)+arrList(13,intLoop),0)%></td>
		<td><%=formatnumber(arrList(7,intLoop)+arrList(16,intLoop),0)%></td>

		<td><%=formatnumber(arrList(2,intLoop)+arrList(11,intLoop),0)%></td>
		<td><%=formatnumber(arrList(5,intLoop)+arrList(14,intLoop),0)%></td>
		<td><%=formatnumber(arrList(8,intLoop)+arrList(17,intLoop),0)%></td>

		<td><%=formatnumber(arrList(3,intLoop)+arrList(12,intLoop),0)%></td>
		<td><%=formatnumber(arrList(6,intLoop)+arrList(15,intLoop),0)%></td>
		<td><%=formatnumber(arrList(9,intLoop)+arrList(18,intLoop),0)%></td>

		<td><%=formatnumber(arrList(12,intLoop),0)%></td>
		<td><%=formatnumber(arrList(15,intLoop),0)%></td>
		<td><%=formatnumber(arrList(18,intLoop),0)%></td>

	  	<td bgcolor="#EEEEEE"><%=formatnumber(arrList(10,intLoop)+arrList(11,intLoop)+arrList(12,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(13,intLoop)+arrList(14,intLoop)+arrList(15,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(arrList(16,intLoop)+arrList(17,intLoop)+arrList(18,intLoop),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((arrList(10,intLoop)+arrList(11,intLoop)+arrList(12,intLoop))-(arrList(13,intLoop)+arrList(14,intLoop)+arrList(15,intLoop)),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((arrList(10,intLoop)+arrList(11,intLoop)+arrList(12,intLoop))-(arrList(16,intLoop)+arrList(17,intLoop)+arrList(18,intLoop)),0)%></td>
	</tr>
	<Tr bgcolor="#FFFFFF" align="right">
	 	<TD  bgcolor="#E6E6E6" align="center">합계</TD>
		<td><%=formatnumber(arrList(1,intLoop)+arrList(10,intLoop),0)%></td>
		<td><%=formatnumber(arrList(4,intLoop)+arrList(13,intLoop),0)%></td>
		<td><%=formatnumber(arrList(7,intLoop)+arrList(16,intLoop),0)%></td>

		<td><%=formatnumber(arrList(2,intLoop)+arrList(11,intLoop),0)%></td>
		<td><%=formatnumber(arrList(5,intLoop)+arrList(14,intLoop),0)%></td>
		<td><%=formatnumber(arrList(8,intLoop)+arrList(17,intLoop),0)%></td>

		<td><%=formatnumber(arrList(3,intLoop)+arrList(12,intLoop),0)%></td>
		<td><%=formatnumber(arrList(6,intLoop)+arrList(15,intLoop),0)%></td>
		<td><%=formatnumber(arrList(9,intLoop)+arrList(18,intLoop),0)%></td>

		<% dim totA, totS, totO
		totA = arrList(1,intLoop)+arrList(10,intLoop)+arrList(2,intLoop)+arrList(11,intLoop)+arrList(3,intLoop)+arrList(12,intLoop)
		totS = arrList(4,intLoop)+arrList(13,intLoop)+arrList(5,intLoop)+arrList(14,intLoop)+arrList(6,intLoop)+arrList(15,intLoop)
		totO = arrList(7,intLoop)+arrList(16,intLoop)+arrList(8,intLoop)+arrList(17,intLoop)+arrList(9,intLoop)+arrList(18,intLoop)
		%>
	  	<td bgcolor="#EEEEEE"><%=formatnumber(totA,0)%></td>
	  	<td bgcolor="#EEEEEE"><%=formatnumber(totS,0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber(totO,0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber(totA-totS,0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber(totA-totO,0)%> </td>
	</tr>
	 <%next


	 %>
	<%else%>
	<tr bgcolor="#ffffff">
		<td colspan="17" align="center">매칭내역이 없습니다.</td>
	</tr>
	<%end if%>
</table>
<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 <Tr bgcolor="#E6E6E6" align="center">
	<TD rowspan="2">제휴몰</TD>
	 <TD rowspan="2" colspan="2">구분</TD>
	 <TD colspan="3"><%=preDate%></TD>
	 <TD colspan="3"><%=nowDate%></TD>
	 <TD colspan="3"><%=nextDate%></TD>

	 <TD colspan="5">TOTAL</TD>
	</tr>
	<Tr bgcolor="#E6E6E6" align="center">
		<TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>

	 <TD>결제(승인)액</TD>
	 	<TD>10x10출고액</TD>
	 	<TD>제휴정산액</TD>
	 	<td>미출고(10x10)</td>
	 	<td>미정산(제휴)</td>
	</tr>
	<tr bgcolor="#FFFFFF" align="right">
		<TD  rowspan="3" align="center"><%=sellsite%></TD>
		<TD  rowspan="3" align="center">계</TD>
		<TD  bgcolor="#E6E6E6" width="50" align="center">상품</TD>
		<td><%=formatnumber(sumItemA(0),0)%></td>
		<td><%=formatnumber(sumItemS(0),0)%></td>
		<td><%=formatnumber(sumItemO(0),0)%></td>

		<td><%=formatnumber(sumItemA(1),0)%></td>
		<td><%=formatnumber(sumItemS(1),0)%></td>
		<td><%=formatnumber(sumItemO(1),0)%></td>

		<td><%=formatnumber(sumItemA(2),0)%></td>
		<td><%=formatnumber(sumItemS(2),0)%></td>
		<td><%=formatnumber(sumItemO(2),0)%></td>

		<td bgcolor="#EEEEEE"><%=formatnumber(sumItemA(0)+sumItemA(1)+sumItemA(2),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(sumItemS(0)+sumItemS(1)+sumItemS(2),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(sumItemO(0)+sumItemO(1)+sumItemO(2),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumItemA(0)+sumItemA(1)+sumItemA(2))-(sumItemS(0)+sumItemS(1)+sumItemS(2)),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumItemA(0)+sumItemA(1)+sumItemA(2))-(sumItemO(0)+sumItemO(1)+sumItemO(2)),0)%></td>
	</tr>
	<Tr bgcolor="#FFFFFF" align="right">
	 	<TD  bgcolor="#E6E6E6" align="center">배송비</TD>
		<td><%=formatnumber(sumDeA(0),0)%></td>
		<td><%=formatnumber(sumDeS(0),0)%></td>
		<td><%=formatnumber(sumDeO(0),0)%></td>

		<td><%=formatnumber(sumDeA(1),0)%></td>
		<td><%=formatnumber(sumDeS(1),0)%></td>
		<td><%=formatnumber(sumDeO(1),0)%></td>

		<td><%=formatnumber(sumDeA(2),0)%></td>
		<td><%=formatnumber(sumDeS(2),0)%></td>
		<td><%=formatnumber(sumDeO(2),0)%></td>

	  	<td bgcolor="#EEEEEE"><%=formatnumber(sumDeA(0)+sumDeA(1)+sumDeA(2),0)%></td>
	  	<td bgcolor="#EEEEEE"><%=formatnumber(sumDeS(0)+sumDeS(1)+sumDeS(2),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber(sumDeO(0)+sumDeO(1)+sumDeO(2),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumDeA(0)+sumDeA(1)+sumDeA(2))-(sumDeS(0)+sumDeS(1)+sumDeS(2)),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumDeA(0)+sumDeA(1)+sumDeA(2))-(sumDeO(0)+sumDeO(1)+sumDeO(2)),0)%> </td>
	</tr>
	<Tr bgcolor="#FFFFFF" align="right">
	 	<TD  bgcolor="#E6E6E6" align="center">합계</TD>
		<td><%=formatnumber(sumItemA(0)+sumDeA(0),0)%></td>
		<td><%=formatnumber(sumItemS(0)+sumDeS(0),0)%></td>
		<td><%=formatnumber(sumItemO(0)+sumDeO(0),0)%></td>

		<td><%=formatnumber(sumItemA(1)+sumDeA(1),0)%></td>
		<td><%=formatnumber(sumItemS(1)+sumDeS(1),0)%></td>
		<td><%=formatnumber(sumItemO(1)+sumDeO(1),0)%></td>

		<td><%=formatnumber(sumItemA(2)+sumDeA(2),0)%></td>
		<td><%=formatnumber(sumItemS(2)+sumDeS(2),0)%></td>
		<td><%=formatnumber(sumItemO(2)+sumDeO(2),0)%></td>

	  	<td bgcolor="#EEEEEE"><%=formatnumber(sumItemA(0)+sumDeA(0)+sumItemA(1)+sumDeA(1)+sumItemA(2)+sumDeA(2),0)%></td>
	  	<td bgcolor="#EEEEEE"><%=formatnumber(sumItemS(0)+sumDeS(0)+sumItemS(1)+sumDeS(1)+sumItemS(2)+sumDeS(2),0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber(sumItemO(0)+sumDeO(0)+sumItemO(1)+sumDeO(1)+sumItemO(2)+sumDeO(2),0)%> </td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumItemA(0)+sumDeA(0)+sumItemA(1)+sumDeA(1)+sumItemA(2)+sumDeA(2))-(sumItemS(0)+sumDeS(0)+sumItemS(1)+sumDeS(1)+sumItemS(2)+sumDeS(2)),0)%></td>
		<td bgcolor="#EEEEEE"><%=formatnumber((sumItemA(0)+sumDeA(0)+sumItemA(1)+sumDeA(1)+sumItemA(2)+sumDeA(2))-(sumItemO(0)+sumDeO(0)+sumItemO(1)+sumDeO(1)+sumItemO(2)+sumDeO(2)),0)%> </td>
	</tr>
	<tr bgcolor="#FFFFFF" align="right">
		<TD   align="center"><%=sellsite%></TD>
		<TD    align="center" colspan="2">선수금잔액</TD>
		<td></td>
		<td></td>
		<td><%=formatnumber((sumItemA(0)+sumDeA(0))-(sumItemO(0)+sumDeO(0)),0)%></td>
		<td></td>
		<td></td>
		<td><%=formatnumber((sumItemA(1)+sumDeA(1))-(sumItemO(1)+sumDeO(1)),0)%></td>
		<td></td>
		<td></td>
		<td><%=formatnumber((sumItemA(2)+sumDeA(2))-(sumItemO(2)+sumDeO(2)),0)%></td>
		<td bgcolor="#EEEEEE"> </td>
	  	<td bgcolor="#EEEEEE"> </td>
		<td bgcolor="#EEEEEE"> </td>
		<td bgcolor="#EEEEEE"> </td>
		<td bgcolor="#EEEEEE"> </td>
	</table>
</div>

<iframe id="ifrDB" name="ifrDB" src="about:blank" frameborder="0" style="width:600;height:400;"></iframe>
<!-- 검색 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->