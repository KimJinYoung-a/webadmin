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
dim sellsite,yyyy1, mm1, ttt
dim nowDate, nextDate1, nextDate2
dim clsJS, arrList, intLoop,arrActMeachul
dim diffdate, i,j, mindate, maxdate
	dim scm1, scm2, scm3, scm4, om1,om2,om3,om4
	dim scm1_d, scm2_d, scm3_d, scm4_d, om1_d,om2_d,om3_d,om4_d
	dim extscm, extom , extscm_d, extom_d
	dim extm, extc, extj
	dim itemmeachul(3), deliverymeachul(3)

   sellsite = requestCheckVar(Request("sellsite"),32)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
  if sellsite ="" then sellsite ="ssg"

   if yyyy1="" then
   	nowDate = left(dateadd("m",-1,date()),7)
   	yyyy1 = left(nowDate,4)
   	mm1=mid(nowDate,6,2)
  else
  	nowDate =yyyy1&"-"&Format00(2,mm1)
	end if

   set clsJS = new CextJungsanMapping
   clsJS.FRectOutMall = sellsite
   clsJS.FRectyyyymm =nowDate
   arrList = clsJS.fnGetextMatchingData
   diffdate = clsJS.Fdiffdate
	mindate = clsJS.Fmindate
	maxdate = clsJS.Fmaxdate
   set clsJS = nothing

	if mindate ="" or isNull(mindate) then mindate = nowDate
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsGoDetail(sellsite,actdt,smdt, omdt,itemdv){
		if (actdt=="미매칭"){
			actdt="N";
			}
		var popItem = window.open("extJungsanMatchingItem.asp?sellsite="+sellsite+"&actdt="+actdt+"&omdt="+omdt+"&smdt="+smdt+"&itemdv="+itemdv,"winItem","");
	}

	function jsExtJungsanDiffMake(){
		if(document.frmDB.sellsite.value ==""){
			alert("제휴몰을 선택해주세요");
			return;
		}

		document.frmDB.target="ifrDB";
		$("#btnSubmit").prop("disabled", true);
		document.frmDB.submit();
	}
</script>
<!-- 검색 시작 -->
<form name="frmDB" method="post" action="http://stscm.10x10.co.kr/admin/maechul/extjungsandata/extjungsanMappingProc.asp">
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
		결제일:
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
		<input type="button" id="btnSubmit" class="button" value="매출매핑(<%= sellsite %>, <%= nowDate %>)" onClick="jsExtJungsanDiffMake('<%= sellsite %>', '<%= nowDate %>');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>
<div width="1000" style="overflow-x:auto;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 <Tr bgcolor="#E6E6E6" align="center">
		<td colspan="3">출고/정산일</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate %>
		<td colspan="3"><%=left(dateadd("m",i,mindate),7)%></td>
		<%next%>
		<td colspan="3">미매칭</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td colspan="3">계</td>
	</tr>
	<Tr bgcolor="#E6E6E6" align="center">
		<td width="50" nowrap>결제일</td>
		<td width="50" nowrap>구분</td>
		<TD width="80" nowrap>결제(승인)액</TD>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate %>
	 	<TD width="80" nowrap>10x10출고액</TD>
		<TD width="80">(10x10)<br/>제휴정산액</TD>
	 	<TD width="80">(제휴)<br/>제휴정산액</TD>
		 <%next %>
		<TD width="80" nowrap>10x10출고액</TD>
	 	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	 	<TD width="80" >(제휴)<br/>제휴정산액</TD>
		<td width="5" bgcolor="#f4f4f4"></td>
		<TD width="80" nowrap>10x10출고액</TD>
	 	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	 	<TD width="80" >(제휴)<br/>제휴정산액</TD>
	</tr>
	<% dim   sumItemA, sumItemS, sumItemO, sumItemOS
	   dim   sumDeA, sumDeS, sumDeO, sumDeOS
	   dim   sumTotS, sumTotO, sumTotOS
	   dim chkDef, chkjsDef, chki

	if isArray(arrList) then
	sumTotS = 0
	sumTotO = 0
	sumTotOS = 0
	chki = -1
		for intLoop = 0 To ubound(arrList,2)
		sumItemA = arrList(1,intLoop)
		sumDeA = arrList(2,intLoop)
		sumItemS = 0
		sumItemO = 0
		sumItemOS = 0
		sumDeS = 0
		sumDeO = 0
		sumDeOS = 0

		 chkDef = False
		if nowdate=arrList(0,intLoop) then chkDef = True
		%>
	<tr bgcolor="<%if chkDef then%>#B2EBF4<%else%>#FFFFFF<%end if%>" align="right">
		<td rowspan="<%if arrList(0,intLoop)="미매칭" or chkDef then%>4<%else%>3<%end if%>"  bgcolor="<%if chkDef then%>#E6E6E6<%else%>#f6f6f6<%end if%>" align="center" width="80"><%=arrList(0,intLoop)%></td>
		<TD  bgcolor="<%if chkDef then%>#E6E6E6<%else%>#f6f6f6<%end if%>" width="50" align="center">상품</TD>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(arrList(1,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%j=0
		for i = 0 To diffdate*2 step 2
		sumItemS = sumItemS +  arrList(7+i,intLoop)
		sumItemO = sumItemO + arrList(7+i+((diffdate+1)*2),intLoop)
		sumItemOS = sumItemOS + arrList(7+i+((diffdate+1)*4),intLoop)
		if left(dateadd("m",j,mindate),7) = nowdate then chki = i
		%>
	 	<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','<%=left(dateadd("m",j,mindate),7)%>','','I')" style="cursor:pointer;"><%=formatnumber(arrList(7+i,intLoop),0)%></span></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','I')" style="cursor:pointer;"><%=formatnumber(arrList(7+i+((diffdate+1)*4),intLoop),0)%></span></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','I')" style="cursor:pointer;"><%=formatnumber(arrList(7+i+((diffdate+1)*2),intLoop),0)%></span></td>
		 <%
		 j=j+1
		 next
		 sumItemS = sumItemS + arrList(3,intLoop)
		 sumItemO= sumItemO + arrList(5,intLoop)
		 sumItemOS = sumItemOS + arrList(7+((diffdate+1)*6),intLoop)
		 %>
		<td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','I')" style="cursor:pointer;"><%=formatnumber(arrList(3,intLoop),0)%></div>
		</td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','I')" style="cursor:pointer;"><%=formatnumber(arrList(7+((diffdate+1)*6),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','I')" style="cursor:pointer;"><%=formatnumber(arrList(5,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%if chkDef then%>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(sumItemS,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(sumItemOS,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(sumItemO,0)%></span></td>
		<%else%>
		<td></td>
		<td></td>
		<td></td>
		<%end if%>
	</tr>
	<tr bgcolor="<%if chkDef then%>#B2EBF4<%else%>#FFFFFF<%end if%>" align="right">
		<TD  bgcolor="<%if chkDef then%>#E6E6E6<%else%>#f6f6f6<%end if%>" width="50" align="center">배송비</TD>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(arrList(2,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%
		 j =0
		for i = 0 To diffdate*2 step 2
		sumDeS = sumDeS +  arrList(8+i,intLoop)
		sumDeO = sumDeO + arrList(8+i+((diffdate+1)*2),intLoop)
		sumDeOS = sumDeOS + arrList(8+i+((diffdate+1)*4),intLoop)
		%>
	 	<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','<%=left(dateadd("m",j,mindate),7)%>','','D')" style="cursor:pointer;"><%=formatnumber(arrList(8+i,intLoop),0)%></span></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','D')" style="cursor:pointer;"><%=formatnumber(arrList(8+i+((diffdate+1)*4),intLoop),0)%></span></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','D')" style="cursor:pointer;"><%=formatnumber(arrList(8+i+((diffdate+1)*2),intLoop),0)%></span></td>
		 <%
		 j = j+1
		 next
		 sumDeS = sumDeS + arrList(4,intLoop)
		 sumDeO = sumDeO + arrList(6,intLoop)
		 sumDeOS = sumDeOS + arrList(8+((diffdate+1)*6),intLoop)
		 %>
		<td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','D')" style="cursor:pointer;"><%=formatnumber(arrList(4,intLoop),0)%></div> </td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','D')" style="cursor:pointer;"><%=formatnumber(arrList(8+((diffdate+1)*6),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','D')" style="cursor:pointer;"><%=formatnumber(arrList(6,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%if chkDef then%>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(sumDeS,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(sumDeOS,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(sumDeO,0)%></span></td>
		<%else%>
		<td></td>
		<td></td>
		<td></td>
		<%end if%>
	</tr>
	<%if chkDef then%>
	<tr bgcolor="#B2EBF4" align="right">
		<TD  bgcolor="#E6E6E6" width="50" align="center">+/-취소액</TD>
		<td>0</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%
		 j =0
		for i = 0 To diffdate*2 step 2
		%>
	 	<td>0</td>
		<td>0</td>
		<td>0</td>
		 <%
		 j = j+1
		 next
 		 sumDeS = sumDeS + arrList(9+((diffdate+1)*6),intLoop)
		 sumDeO = sumDeO + arrList(10+((diffdate+1)*6),intLoop)
		 sumDeOS = sumDeOS + arrList(10+((diffdate+1)*6),intLoop)
		 %>
		<td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','D')" style="cursor:pointer;"><%=formatnumber(arrList(9+((diffdate+1)*6),intLoop),0)%></div></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','D')" style="cursor:pointer;"><%=formatnumber(arrList(10+((diffdate+1)*6),intLoop),0)%></span></td>
		<td>0</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','D')" style="cursor:pointer;"><%=formatnumber(arrList(9+((diffdate+1)*6),intLoop),0)%></div></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','D')" style="cursor:pointer;"><%=formatnumber(arrList(10+((diffdate+1)*6),intLoop),0)%></span></td>
		<td>0</td>
	</tr>
	<%end if%>
	<%if arrList(0,intLoop) ="미매칭" then%>
	<tr bgcolor="<%if chkDef then%>#B2EBF4<%else%>#FFFFFF<%end if%>" align="right">
		<TD  bgcolor="<%if chkDef then%>#E6E6E6<%else%>#f6f6f6<%end if%>" width="50" align="center">+/-취소액</TD>
		<td>0</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<% j = 0
		for i = 0 To diffdate*2 step 2%>
	 	<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>>
		<%if  i= chki  then%>
		<span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','S')" style="cursor:pointer;<%if chkDef then%>font-weight:bold;<%end if%>">
		<%=formatnumber(arrList(11+((diffdate+1)*6),intLoop),0)%>
		</span>
		<%end if%>
		</td>
		 <%
		  j=j+1
		 next
		 sumDeO = sumDeO + arrList(11+((diffdate+1)*6),intLoop)
		 %>
		 <td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','S')" style="cursor:pointer;"></div>
		 </td>
		 <td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','S')" style="cursor:pointer;"> </td>
		 <td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','S')" style="cursor:pointer;"></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<%end if%>
	<tr bgcolor="<%if chkDef then%>#B2EBF4<%else%>#FFFFFF<%end if%>" align="right">
		<TD  bgcolor="<%if chkDef then%>#E6E6E6<%else%>#f6f6f6<%end if%>" width="50" align="center">계</TD>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','S')" style="cursor:pointer;<%if chkDef then%>font-weight:bold;<%end if%>"><%=formatnumber(arrList(1,intLoop)+arrList(2,intLoop),0)%></b></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<% j = 0
		for i = 0 To diffdate*2 step 2%>
	 	<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','<%=left(dateadd("m",j,mindate),7)%>','','S')" style="cursor:pointer;<%if chkDef then%>font-weight:bold;<%end if%>"><%=formatnumber(arrList(7+i,intLoop)+arrList(8+i,intLoop),0)%></span></td>
		 <td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','S')" style="cursor:pointer;<%if chkDef then%>font-weight:bold;<%end if%>"><%=formatnumber(arrList(7+i+((diffdate+1)*4),intLoop)+arrList(8+i+((diffdate+1)*4),intLoop),0)%></span></td>
		<td <%if i= chki and not chkDef then %>bgcolor="#E6FFFF"<%end if%>><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','<%=left(dateadd("m",j,mindate),7)%>','S')" style="cursor:pointer;<%if chkDef then%>font-weight:bold;<%end if%>">
		<%if arrList(0,intLoop) ="미매칭" then%>
		<%=formatnumber(arrList(7+i+((diffdate+1)*2),intLoop)+arrList(8+i+((diffdate+1)*2),intLoop)+arrList(11+((diffdate+1)*6),intLoop),0)%>
		<%else%>
		<%=formatnumber(arrList(7+i+((diffdate+1)*2),intLoop)+arrList(8+i+((diffdate+1)*2),intLoop),0)%>
		<%end if %>
		</span></td>
		 <%
		  j=j+1
		 next %>
		 <td><div onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','N','','S')" style="cursor:pointer;"><%=formatnumber(arrList(3,intLoop)+arrList(4,intLoop)+arrList(9+((diffdate+1)*6),intLoop),0)%></div> </td>
		 <td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','S')" style="cursor:pointer;"><%=formatnumber(arrList(7+((diffdate+1)*6),intLoop)+arrList(8+((diffdate+1)*6),intLoop)+arrList(10+((diffdate+1)*6),intLoop),0)%></span></td>
		 <td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','N','S')" style="cursor:pointer;"><%=formatnumber(arrList(5,intLoop)+arrList(6,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%if chkDef then%>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','S')" style="cursor:pointer;"><b><%=formatnumber(sumItemS + sumDeS,0)%></b></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','S')" style="cursor:pointer;"><b><%=formatnumber(sumItemOS + sumDeOS,0)%></b></span></td>
		<td><span onClick="jsGoDetail('<%=sellsite%>','<%=arrList(0,intLoop)%>','','','S')" style="cursor:pointer;"><b><%=formatnumber(sumItemO + sumDeO,0)%></b></span></td>
		<%else%>
		<td></td>
		<td></td>
		<td></td>
		<%end if%>
	</tr>

	 <%
	 sumTotS = sumTotS + arrList(7+chki,intLoop)+arrList(8+chki,intLoop)
	 if arrList(0,intLoop) ="미매칭" then
	 sumTotO = sumTotO + (arrList(7+chki+((diffdate+1)*2),intLoop) + arrList(8+chki+((diffdate+1)*2),intLoop)+arrList(11+((diffdate+1)*6),intLoop))
	 else
	 sumTotO = sumTotO + (arrList(7+chki+((diffdate+1)*2),intLoop) + arrList(8+chki+((diffdate+1)*2),intLoop))
	 end if
	 sumTotOS = sumTotOS + ( arrList(7+chki+((diffdate+1)*4),intLoop)+ arrList(8+chki+((diffdate+1)*4),intLoop))

	 next%>
	 <tr bgcolor="#ffffff" align="right">
	 	<td colspan="3" align="center" bgcolor="#E6E6E6">합계</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%if chki >0 then%><td colspan="<%=(chki/2)*3%>"></td><%end if%>
		<td bgcolor="#E6FFFF"><b><%=formatnumber(sumTotS,0)%></b></td>
		<td bgcolor="#E6FFFF"><b><%=formatnumber(sumTotOS,0)%></b></td>
		<td bgcolor="#E6FFFF"><b><%=formatnumber(sumTotO,0)%></b></td>
		 <Td colspan="<%if chki >0 then%><%=(diffdate-(chki/2)+1)*3%><%else%><%=(diffdate+1)*3%><%end if%>"></td>
		 <td width="5" bgcolor="#f4f4f4"></td>
		 <td colspan="3"></td>
	 </tr>
	<%else%>
	<tr bgcolor="#ffffff">
		<td colspan="21" align="center">매칭내역이 없습니다.</td>
	</tr>
	<%end if%>

	</table>
</div>

<iframe id="ifrDB" name="ifrDB" src="about:blank" frameborder="0" style="width:600;height:400;"></iframe>
<!-- 검색 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->