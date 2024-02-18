<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsanDiffcls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sellsite,yyyy1, mm1 
dim nowDate, preDate, nextDate
dim clsJS, arrList, intLoop
   sellsite = requestCheckVar(Request("sellsite"),10)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
  if sellsite ="" then sellsite ="ssg"
   'yyyy1 = "2018"
   'mm1 = "03"
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
   clsJS.FRectPyyyymm =preDate
   clsJS.FRectNyyyymm =nextDate
   arrList = clsJS.fnGetextMappingData   
   set clsJS = nothing
   
%>
<script type="text/javascript">
	function jsGoDetail(sellsite,stype, yyyy1,mm1,yyyy2,mm2){
		var popItem = window.open("extJungsanMappingItem.asp?sellsite="+sellsite+"&yyyy1="+yyyy1+"&mm1="+mm1+"&yyyy2="+yyyy2+"&mm2="+mm2+"&stype="+stype,"winItem","");
	}
</script>
<!-- 검색 시작 -->
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

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 
	<Tr bgcolor="#ffffff" align="center">  
		<td rowspan="2" colspan="3" ></td>   
		<td colspan="3" width="15%">scm <%=preDate%></td>
		<td colspan="3" width="15%">scm <%=nowDate%></td> 
		<td colspan="3" width="15%">scm 미매칭</td>
		<td colspan="6" bgcolor="#eeeeee"><b>scm 합계</b></td>
	</tr>	
	<Tr bgcolor="#ffffff" align="center">     
		<td>상품</td>
		<td>배송비</td>
		<td>합계</td>
		<td>상품</td>
		<td>배송비</td>
		<td>합계</td> 
		<td>상품</td>
		<td>배송비</td>
		<td>합계</td>
		<td bgcolor="#eeeeee"><b>상품</b></td>
		<td bgcolor="#eeeeee"><b>배송비</b></td>
		<td bgcolor="#eeeeee"><b>총합</b></td>
		<td bgcolor="#eeeeee"><b>제휴매출금액</b></td>
		<td bgcolor="#eeeeee"><b>제휴수수료</b></td>
		<td bgcolor="#eeeeee"><b>제휴정산금액</b></td>
	</tr>	
	<%
	dim scm1, scm2, scm3, scm4, om1,om2,om3,om4
	dim scm1_d, scm2_d, scm3_d, scm4_d, om1_d,om2_d,om3_d,om4_d
	dim extscm, extom , extscm_d, extom_d
	if isArray(arrList) then
		scm1=0:		scm2=0:		scm3=0:		scm4=0
		om1=0: om2=0: om3=0:om4=0
		scm1_d=0:		scm2_d=0:		scm3_d=0:		scm4_d=0
		om1_d=0: om2_d=0: om3_d=0:om4_d=0
	 
		%>
	<% for intLoop =0 To uBound(arrList,2)%>
	<tr bgcolor="#e3f1fb" align="right">
		<%if intLoop =0 then%>
		<td rowspan="10" align="center" width="50" bgcolor="ffdddd"><%=sellsite%></td>
		<%end if%>
		<td bgcolor="#FFFFFF" rowspan="2" align="center" width="100"><%=arrList(0,intLoop)%></td>
		<td align="center" width="100">scm</td>
		<td ><a href="javascript:jsGoDetail('<%=sellsite%>','I','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(1,intLoop)<>arrList(5,intLoop) then%>style="color:red" <%end if%>><%=formatnumber(arrList(1,intLoop),0)%></span></a>
		</td>
		<td ><a href="javascript:jsGoDetail('<%=sellsite%>','D','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(9,intLoop)<>arrList(13,intLoop) then%>style="color:red"<%end if%>><%=formatnumber(arrList(9,intLoop),0)%></span></a>
		</td>	
		<td ><a href="javascript:jsGoDetail('<%=sellsite%>','A','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
		<span <%if (arrList(1,intLoop)+arrList(9,intLoop))<>(arrList(5,intLoop)+arrList(13,intLoop)) then%>style="color:red"<%end if%>><%=formatnumber( (arrList(1,intLoop)+arrList(9,intLoop)),0)%></span></a>
		</td>	
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','I','<%=left(nowDate,4)%>','<%=mid(nowDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(2,intLoop)<>arrList(6,intLoop) then%> style="color:red"<%end if%>><%=formatnumber(arrList(2,intLoop),0)%></span></a>
		</td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','D','<%=left(nowDate,4)%>','<%=mid(nowDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(10,intLoop)<>arrList(14,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(10,intLoop),0)%></span></a>
		</td>
		<td ><a href="javascript:jsGoDetail('<%=sellsite%>','A','<%=left(nowDate,4)%>','<%=mid(nowDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if (arrList(2,intLoop)+arrList(10,intLoop))<>(arrList(6,intLoop)+arrList(14,intLoop)) then%>style ="color:red"<%end if%>><%=formatnumber( (arrList(2,intLoop)+arrList(10,intLoop)),0)%></span></a>
		</td>	 
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','I','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(4,intLoop)<>arrList(8,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(4,intLoop),0)%></span></a>
		</td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','D','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(12,intLoop)<>arrList(16,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(12,intLoop),0)%></span></a>
		</td>
		<td ><a href="javascript:jsGoDetail('<%=sellsite%>','A','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if (arrList(4,intLoop)+arrList(12,intLoop))<>(arrList(8,intLoop)+arrList(16,intLoop)) then%>style ="color:red"<%end if%>><%=formatnumber( (arrList(4,intLoop)+arrList(12,intLoop)),0)%></span></a>
		</td>
		 	<% extscm = 0 : extscm_d = 0 : extom =0 : extom_d =0
		 	extscm = arrList(1,intLoop)+arrList(2,intLoop)+arrList(3,intLoop)+arrList(4,intLoop)
		 	extscm_d = arrList(9,intLoop)+arrList(10,intLoop)+arrList(11,intLoop)+arrList(12,intLoop)
		 	extom = arrList(5,intLoop)+arrList(6,intLoop)+arrList(7,intLoop)+arrList(8,intLoop)
		 	extom_d = arrList(13,intLoop)+arrList(14,intLoop)+arrList(15,intLoop)+arrList(16,intLoop)
		 	%>
		<td><span <%if  (extscm<>extom) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extscm,0)%></strong></span>
		</td>
		<td><span <%if (extscm_d<>extom_d) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extscm_d,0)%></strong></span>
		</td>
		<td><span <%if (extscm+ extscm_d)<>( extom+extom_d) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extscm+ extscm_d,0)%></strong></span>
		</td>
	</tr>
	<tr bgcolor="#ffDDDD" align="right">   
		<td align="center" width="100"><%=sellsite%></td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','I','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(1,intLoop)<>arrList(5,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(5,intLoop),0)%></span></a>
		</td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','D','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(9,intLoop)<>arrList(13,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(13,intLoop),0)%></span></a>
		</td>
	  <td>
	  	<a href="javascript:jsGoDetail('<%=sellsite%>','A','<%=left(preDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if (arrList(1,intLoop)+arrList(9,intLoop))<>(arrList(5,intLoop)+arrList(13,intLoop)) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(5,intLoop)+arrList(13,intLoop),0)%></span></a>
	  </td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','I','<%=left(nowDate,4)%>','<%=mid(nowDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(2,intLoop)<>arrList(6,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(6,intLoop),0)%></span></a>
		</td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','D','<%=left(nowDate,4)%>','<%=mid(nowDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(10,intLoop)<>arrList(14,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(14,intLoop),0)%></span></a>
		</td>
		<td>
			<a href="javascript:jsGoDetail('<%=sellsite%>','A','<%=left(nowDate,4)%>','<%=mid(preDate,6,2)%>','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if (arrList(2,intLoop)+arrList(10,intLoop))<>(arrList(6,intLoop)+arrList(14,intLoop)) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(6,intLoop)+arrList(14,intLoop),0)%></span></a>
		</td> 
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','I','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(4,intLoop)<>arrList(8,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(8,intLoop),0)%></span></a>
		</td>
		<td><a href="javascript:jsGoDetail('<%=sellsite%>','D','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if arrList(12,intLoop)<>arrList(16,intLoop) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(16,intLoop),0)%></span></a>
		</td>
		<td>
			<a href="javascript:jsGoDetail('<%=sellsite%>','A','','','<%=left(arrList(0,intLoop),4)%>','<%=mid(arrList(0,intLoop),6,2)%>');">
			<span <%if (arrList(4,intLoop)+arrList(12,intLoop))<>(arrList(8,intLoop)+arrList(16,intLoop)) then%>style ="color:red"<%end if%>><%=formatnumber(arrList(8,intLoop)+arrList(16,intLoop),0)%></span></a>
		</td>
			<td><span <%if  (extscm<>extom) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extom,0)%></strong></span>
		</td>
		<td><span <%if (extscm_d<>extom_d) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extom_d,0)%></strong></span>
		</td>
		<td><span <%if (extscm+ extscm_d)<>( extom+extom_d) then%>style="color:red"<%end if%>>
			<strong><%=formatnumber(extom+ extom_d,0)%></strong></span>
		</td>
		<td><%=formatnumber(arrList(17,intLoop),0)%></td>
		<td><%=formatnumber(arrList(18,intLoop),0)%></td>
		<td><%=formatnumber(arrList(19,intLoop),0)%></td>
	</tr>
	<% scm1 = scm1 + arrList(1,intLoop) 
	scm2 = scm2 + arrList(2,intLoop) 
	scm3 = scm3 + arrList(3,intLoop) 
	scm4 = scm4 + arrList(4,intLoop) 
	om1 = om1+ arrList(5,intLoop) 
	om2 = om2+ arrList(6,intLoop) 
	om3 = om3+ arrList(7,intLoop) 
	om4 = om4+ arrList(8,intLoop) 
	scm1_d = scm1_d + arrList(9,intLoop) 
	scm2_d = scm2_d + arrList(10,intLoop) 
	scm3_d = scm3_d + arrList(11,intLoop) 
	scm4_d = scm4_d + arrList(12,intLoop) 
	om1_d = om1_d+ arrList(13,intLoop) 
	om2_d = om2_d+ arrList(14,intLoop) 
	om3_d = om3_d+ arrList(15,intLoop) 
	om4_d = om4_d+ arrList(16,intLoop)
	%>
	<% next%>
	<tr bgcolor="#e3f1fb" align="right">
		<td rowspan="2" bgcolor="#eeeeee" align="center">합계</td>
		<td align="center">scm</td>
		<td><span <%if scm1 <> om1 then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1,0)%></strong></span></td>
		<td><span <%if scm1_d <> om1_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1_d,0)%></strong></span></td>
		<td><span <%if (scm1+scm1_d) <> (om1+om1_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1+scm1_d,0)%></strong></span></td>	
		<td><span <%if scm2 <> om2 then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm2,0)%></strong></span></td>
		<td><span <%if scm2_d <> om2_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm2_d,0)%></strong></span></td>
		<td><span <%if (scm2+scm2_d) <> (om2+om2_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm2+scm2_d,0)%></strong></span></td>
		<td><span <%if scm3 <> om3 then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm3,0)%></strong></span></td>
		<td><span <%if scm3_d <> om3_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm3_d,0)%></strong></span></td>
		<td><span <%if (scm3+scm3_d) <> (om3+om3_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm3+scm3_d,0)%></strong></span></td>		 
		<td><span <%if (scm1+scm2+scm3+scm4) <> (om1+om2+om3+om4) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1+scm2+scm3+scm4,0)%></strong></span></td>
		<td><span <%if (scm1_d+scm2_d+scm3_d+scm4_d) <> (om1_d+om2_d+om3_d+om4_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1_d+scm2_d+scm3_d+scm4_d,0)%></strong></span></td>
		<td><span <%if (scm1+scm2+scm3+scm4+scm1_d+scm2_d+scm3_d+scm4_d) <> (om1+om2+om3+om4+om1_d+om2_d+om3_d+om4_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(scm1+scm2+scm3+scm4+scm1_d+scm2_d+scm3_d+scm4_d,0)%></strong></span></td>
	</tr>
	<tr  bgcolor="ffdddd" align="right">
		<td align="center"><%=sellsite%></td>   
		<td><span <%if scm1 <> om1 then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1,0)%></strong></span></td>
		<td><span <%if scm1_d <> om1_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1_d,0)%></strong></span></td>
		<td><span <%if (scm1+scm1_d) <> (om1+om1_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1+om1_d,0)%></strong></span></td>	
		<td><span <%if scm2 <> om2 then%>style ="color:red"<%end if%>><strong><%=formatnumber(om2,0)%></strong></span></td>
		<td><span <%if scm2_d <> om2_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(om2_d,0)%></strong></span></td>
		<td><span <%if (scm2+scm2_d) <> (om2+om2_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om2+om2_d,0)%></strong></span></td>	
		<td><span <%if scm3 <> om3 then%>style ="color:red"<%end if%>><strong><%=formatnumber(om3,0)%></strong></span></td>
		<td><span <%if scm3_d <> om3_d then%>style ="color:red"<%end if%>><strong><%=formatnumber(om3_d,0)%></strong></span></td>
		<td><span <%if (scm3+scm3_d) <> (om3+om3_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om3+om3_d,0)%></strong></span></td>	 
		<td><span <%if (scm1+scm2+scm3+scm4) <> (om1+om2+om3+om4) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1+om2+om3+om4,0)%></strong></span></td>
		<td><span <%if (scm1_d+scm2_d+scm3_d+scm4_d) <> (om1_d+om2_d+om3_d+om4_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1_d+om2_d+om3_d+om4_d,0)%></strong></span></td>
		<td><span <%if (scm1+scm2+scm3+scm4+scm1_d+scm2_d+scm3_d+scm4_d) <> (om1+om2+om3+om4+om1_d+om2_d+om3_d+om4_d) then%>style ="color:red"<%end if%>><strong><%=formatnumber(om1+om2+om3+om4+om1_d+om2_d+om3_d+om4_d,0)%></strong></span></td>
	</tr>
	<%else%>
	<tr>
		<td colspan="13" align="center">매칭내역이 없습니다.</td>
	</tr>
	<%end if%>
</table>
<!-- 검색 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->