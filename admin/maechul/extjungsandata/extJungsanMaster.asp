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
dim sellsite,yyyy1, mm1, yyyy2, mm2
dim stDate, edDate
dim clsJS, arrList, intLoop
dim diffdate, i,j, mindate, maxdate

   sellsite = requestCheckVar(Request("sellsite"),32)
   yyyy1 = requestCheckVar(Request("yyyy1"),4)
   mm1 = requestCheckVar(Request("mm1"),2)
   yyyy2 = requestCheckVar(Request("yyyy2"),4)
   mm2 = requestCheckVar(Request("mm2"),2)
  'if sellsite ="" then sellsite ="ssg"

   if yyyy1="" then
   	stDate = left(dateadd("m",-1,date()),7)
   	yyyy1 = left(stDate,4)
   	mm1=mid(stDate,6,2)
    else
  	stDate =yyyy1&"-"&Format00(2,mm1)
	end if

    if yyyy2="" then
   	edDate = left(dateadd("m",-1,date()),7)
   	yyyy2 = left(edDate,4)
   	mm2=mid(edDate,6,2)
    else
  	edDate =yyyy2&"-"&Format00(2,mm2)
	end if

   set clsJS = new CextJungsanMapping
   clsJS.FRectOutMall = sellsite
   clsJS.FRectStyyyymm =stDate
   clsJS.FRectedyyyymm =edDate
   arrList = clsJS.fnGetextMatchingMaster
   diffdate = clsJS.Fdiffdate
   mindate = clsJS.Fmindate
   maxdate = clsJS.Fmaxdate

   '// set clsJS = nothing

   clsJS.fnGetextMatchingMaster_V2()
   '// dbget.close() : response.end

dim currYYYYMM
dim payPrice
dim ten_meachul_yyyymm0
dim ten_deliver_meachul_yyyymm0
dim ten_jungsan_yyyymm0
dim ext_jungsan_yyyymm0

dim ten_meachul_yyyymm1
dim ten_deliver_meachul_yyyymm1
dim ten_jungsan_yyyymm1
dim ext_jungsan_yyyymm1

dim ten_meachul_yyyymm2
dim ten_deliver_meachul_yyyymm2
dim ten_jungsan_yyyymm2
dim ext_jungsan_yyyymm2

dim ten_meachul_yyyymm3
dim ten_deliver_meachul_yyyymm3
dim ten_jungsan_yyyymm3
dim ext_jungsan_yyyymm3

dim ten_meachul_null
dim ten_deliver_meachul_null
dim ten_jungsan_null
dim ext_jungsan_null

dim ten_meachul_sum
dim ten_deliver_meachul_sum
dim ten_jungsan_sum
dim ext_jungsan_sum

dim tot_ten_meachul_yyyymm0
dim tot_ten_deliver_meachul_yyyymm0
dim tot_ten_jungsan_yyyymm0
dim tot_ext_jungsan_yyyymm0

dim tot_ten_meachul_yyyymm1
dim tot_ten_deliver_meachul_yyyymm1
dim tot_ten_jungsan_yyyymm1
dim tot_ext_jungsan_yyyymm1

dim tot_ten_meachul_yyyymm2
dim tot_ten_deliver_meachul_yyyymm2
dim tot_ten_jungsan_yyyymm2
dim tot_ext_jungsan_yyyymm2

dim tot_ten_meachul_yyyymm3
dim tot_ten_deliver_meachul_yyyymm3
dim tot_ten_jungsan_yyyymm3
dim tot_ext_jungsan_yyyymm3

dim tot_ten_meachul_null
dim tot_ten_deliver_meachul_null
dim tot_ten_jungsan_null
dim tot_ext_jungsan_null

dim tot_ten_meachul_sum
dim tot_ten_deliver_meachul_sum
dim tot_ten_jungsan_sum
dim tot_ext_jungsan_sum

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
	function jsGoDetail(sellsite,actdt,smdt, omdt,itemdv, scmdeliverdate){
		if (actdt=="미매칭"){
			actdt="N";
		}
        if (scmdeliverdate == undefined) {
            scmdeliverdate = ""
        }
		var popItem = window.open("extJungsanMatchingItem.asp?sellsite="+sellsite+"&actdt="+actdt+"&omdt="+omdt+"&smdt="+smdt+"&itemdv="+itemdv+"&scmdeliverdate="+scmdeliverdate,"winItem","");
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

	function popjsExtJungsanDiffMake(sellsite, jsdate){
		if(document.frmDB.sellsite.value ==""){
			alert("제휴몰을 선택해주세요");
			return;
		}
		$("#btnSubmit").prop("disabled", true);
		<% If application("Svr_Info") <> "Dev" Then %>
			var popwin = window.open("https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extjungsanMappingProc.asp?sellsite="+sellsite+"&jsdate="+jsdate,"popjsExtJungsanDiffMake","width=600,height=300,scrollbars=yes,resizable=yes");
		<% Else %>
			var popwin = window.open("/admin/maechul/extjungsandata/extjungsanMappingProc.asp?sellsite="+sellsite+"&jsdate="+jsdate,"popjsExtJungsanDiffMake","width=600,height=300,scrollbars=yes,resizable=yes");
		<% End If %>
		popwin.focus();
	}

</script>
<!-- 검색 시작 -->
<% If application("Svr_Info") <> "Dev" Then %>
<form name="frmDB" method="post" action="https://stscm.10x10.co.kr/admin/maechul/extjungsandata/extjungsanMappingProc.asp">
<% Else %>
<form name="frmDB" method="post" action="/admin/maechul/extjungsandata/extjungsanMappingProc.asp">
<% End If %>
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sellsite" value="<%= sellsite %>">
<input type="hidden" name="jsdate" value="<%=eddate%>">
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
		제휴몰:	<% 'fnGetOptOutMall sellsite %>
				<% call drawOutmallSelectBox("sellsite",sellsite) %>
		&nbsp;
		결제일:
		<% DrawYMYMBox yyyy1,mm1 , yyyy2,mm2%>
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
		<input type="button" id="btnSubmit" class="button" value="매출매핑(<%= sellsite %>, <%= edDate %>)" onClick="popjsExtJungsanDiffMake('<%= sellsite %>', '<%= edDate %>');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p />

[NEW]
<div width="900" style="overflow-x:auto;overflow-y:hidden;">
<table    align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<Tr bgcolor="#E6E6E6" align="center">
	<td colspan="4">출고/정산일</td>
	<td width="5" bgcolor="#f4f4f4"></td>
    <td colspan="4"><%=left(dateadd("m", 0, mindate), 7)%></td>
    <td colspan="4"><%=left(dateadd("m", 1, mindate), 7)%></td>
    <% if (diffdate >= 2) then %>
    <td colspan="4"><%=left(dateadd("m", 2, mindate), 7)%></td>
    <% end if %>
    <% if (diffdate >= 3) then %>
    <td colspan="4"><%=left(dateadd("m", 3, mindate), 7)%></td>
    <% end if %>
	<td colspan="3">미매칭</td>
	<td width="5" bgcolor="#f4f4f4"></td>
	<td colspan="4">합계</td>
</tr>
<Tr bgcolor="#E6E6E6" align="center">
	<td width="50" nowrap>결제일</td>
	<td width="50" nowrap>제휴몰</td>
	<td width="50" nowrap>구분</td>
	<TD width="80" nowrap>결제(승인)액</TD>
	<td width="5" bgcolor="#f4f4f4"></td>
	<TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80">(10x10)<br/>제휴정산액</TD>
	<TD width="80">(제휴)<br/>제휴정산액</TD>
	<TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80">(10x10)<br/>제휴정산액</TD>
	<TD width="80">(제휴)<br/>제휴정산액</TD>
    <% if (diffdate >= 2) then %>
    <TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80">(10x10)<br/>제휴정산액</TD>
	<TD width="80">(제휴)<br/>제휴정산액</TD>
    <% end if %>
    <% if (diffdate >= 3) then %>
    <TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80">(10x10)<br/>제휴정산액</TD>
	<TD width="80">(제휴)<br/>제휴정산액</TD>
    <% end if %>
	<TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	<!--<TD width="80" >(제휴)<br/>제휴정산액</TD>-->
	<td width="5" bgcolor="#f4f4f4"></td>

	<TD width="80" nowrap>10x10<br />출고액</TD>
    <TD width="80" nowrap>10x10<br />정산액</TD>
	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	<TD width="80" >(제휴)<br/>제휴정산액</TD>
</tr>
<% currYYYYMM = "" %>
<% for i = 0 to clsJS.FResultCount - 1 %>
<% if (i mod 3) = 0 then %>
	<tr bgcolor="#FFFFFF" align="right">
		<td rowspan="4"  width="50" bgcolor="#f6f6f6" align="center"><%= clsJS.FItemList(i).Fyyyymm %></td>
		<td rowspan="4"  width="50" bgcolor="#f6f6f6" align="center"><%= clsJS.FItemList(i).Fsitename %></td>
		<td nowrap  width="50" bgcolor="#f6f6f6" align="center">상품</td>
		<td width="80"><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;<%if clsJS.FItemList(i).Fyyyymm < stDate and  clsJS.FItemList(i).Fyyyymm > edDate then%>color:gray;<%end if%>"><%= FormatNumber(clsJS.FItemList(i).FpayPrice, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>

        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",0,mindate)%>','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '<%=dateadd("m",0,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Fext_jungsan_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",1,mindate)%>','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_yyyymm1, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '<%=dateadd("m",1,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Fext_jungsan_yyyymm1, 0) %></span></td>

        <% if (diffdate >= 2) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",2,mindate)%>','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_yyyymm2, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '<%=dateadd("m",2,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Fext_jungsan_yyyymm2, 0) %></span></td>
        <% end if %>
        <% if (diffdate >= 3) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",3,mindate)%>','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_yyyymm3, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '<%=dateadd("m",3,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Fext_jungsan_yyyymm3, 0) %></span></td>
        <% end if %>

		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','N','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_null, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', 'N')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_null, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','N','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_null, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_meachul_sum, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_deliver_meachul_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Ften_jungsan_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i).Fext_jungsan_sum, 0) %></span></td>
	</tr>

	<tr bgcolor="#FFFFFF" align="right">
		<td nowrap width="50" bgcolor="#f6f6f6" align="center">배송비</td>
		<td width="80"><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '')" style="cursor:pointer;<%if clsJS.FItemList(i+1).Fyyyymm < stDate and  clsJS.FItemList(i+1).Fyyyymm > edDate then%>color:gray;<%end if%>"><%= FormatNumber(clsJS.FItemList(i+1).FpayPrice, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>

        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','<%=dateadd("m",0,mindate)%>','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '<%=dateadd("m",0,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Fext_jungsan_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','<%=dateadd("m",1,mindate)%>','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_yyyymm1, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '<%=dateadd("m",1,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Fext_jungsan_yyyymm1, 0) %></span></td>

        <% if (diffdate >= 2) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','<%=dateadd("m",2,mindate)%>','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_yyyymm2, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '<%=dateadd("m",2,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Fext_jungsan_yyyymm2, 0) %></span></td>
        <% end if %>
        <% if (diffdate >= 3) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','<%=dateadd("m",3,mindate)%>','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_yyyymm3, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '<%=dateadd("m",3,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Fext_jungsan_yyyymm3, 0) %></span></td>
        <% end if %>

        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','N','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_null, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', 'N')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_null, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','N','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_null, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_meachul_sum, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_deliver_meachul_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Ften_jungsan_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+1).Fsitename %>','<%= clsJS.FItemList(i+1).Fyyyymm %>','','','D', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+1).Fext_jungsan_sum, 0) %></span></td>
	</tr>

	<tr bgcolor="#FFFFFF" align="right">
		<td nowrap bgcolor="#f6f6f6" align="center">+/-취소액
		    <%if clsJS.FItemList(i+2).Fyyyymm >= stDate and clsJS.FItemList(i+2).Fyyyymm <= edDate   then%>
		    <div style="color:gray;">[미매칭취소액]</div>
		    <%end if%>
		</td>
		<td width="80"><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '')" style="cursor:pointer;<%if clsJS.FItemList(i+2).Fyyyymm < stDate and  clsJS.FItemList(i+2).Fyyyymm > edDate then%>color:gray;<%end if%>"><%= FormatNumber(clsJS.FItemList(i+2).FpayPrice, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>

        <td>
            <span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','<%=dateadd("m",0,mindate)%>','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_meachul_yyyymm0, 0) %></span>
            <%if clsJS.FItemList(i+2).Ften_meachul_null <> 0 and clsJS.FItemList(i+2).Fyyyymm = left(dateadd("m",0,mindate),7) then%><div onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','N','','M', '')" style="cursor:pointer;color:gray">[<%=formatnumber(clsJS.FItemList(i+2).Ften_meachul_null,0)%>]</div><%end if%>
        </td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '<%=dateadd("m",0,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm0, 0) %></span></td>
		<td>
            <span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_jungsan_yyyymm0, 0) %></span>
            <%if clsJS.FItemList(i+2).Ften_jungsan_null <> 0 and clsJS.FItemList(i+2).Fyyyymm = left(dateadd("m",0,mindate),7)  then%><div onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','N','M', '')" style="cursor:pointer;color:gray">[<%=formatnumber(clsJS.FItemList(i+2).Ften_jungsan_null,0)%>]</div><%end if%>
        </td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Fext_jungsan_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','<%=dateadd("m",1,mindate)%>','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_meachul_yyyymm1, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '<%=dateadd("m",1,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_jungsan_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Fext_jungsan_yyyymm1, 0) %></span></td>

        <% if (diffdate >= 2) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','<%=dateadd("m",2,mindate)%>','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_meachul_yyyymm2, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '<%=dateadd("m",2,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_jungsan_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Fext_jungsan_yyyymm2, 0) %></span></td>
        <% end if %>
        <% if (diffdate >= 3) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','<%=dateadd("m",3,mindate)%>','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_meachul_yyyymm3, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '<%=dateadd("m",3,mindate)%>')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_jungsan_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Fext_jungsan_yyyymm3, 0) %></span></td>
        <% end if %>

        <td></td>
        <td></td>
		<td></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_meachul_sum, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_deliver_meachul_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Ften_jungsan_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i+2).Fsitename %>','<%= clsJS.FItemList(i+2).Fyyyymm %>','','','M', '')" style="cursor:pointer;"><%= formatnumber(clsJS.FItemList(i+2).Fext_jungsan_sum, 0) %></span></td>
	</tr>

    <%
    payPrice = clsJS.FItemList(i).FpayPrice + clsJS.FItemList(i+1).FpayPrice + clsJS.FItemList(i+2).FpayPrice

    ten_meachul_yyyymm0 = clsJS.FItemList(i).Ften_meachul_yyyymm0 + clsJS.FItemList(i+1).Ften_meachul_yyyymm0 + clsJS.FItemList(i+2).Ften_meachul_yyyymm0
    ten_deliver_meachul_yyyymm0 = clsJS.FItemList(i).Ften_deliver_meachul_yyyymm0 + clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm0 + clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm0
    ten_jungsan_yyyymm0 = clsJS.FItemList(i).Ften_jungsan_yyyymm0 + clsJS.FItemList(i+1).Ften_jungsan_yyyymm0 + clsJS.FItemList(i+2).Ften_jungsan_yyyymm0
    ext_jungsan_yyyymm0 = clsJS.FItemList(i).Fext_jungsan_yyyymm0 + clsJS.FItemList(i+1).Fext_jungsan_yyyymm0 + clsJS.FItemList(i+2).Fext_jungsan_yyyymm0

    ten_meachul_yyyymm1 = clsJS.FItemList(i).Ften_meachul_yyyymm1 + clsJS.FItemList(i+1).Ften_meachul_yyyymm1 + clsJS.FItemList(i+2).Ften_meachul_yyyymm1
    ten_deliver_meachul_yyyymm1 = clsJS.FItemList(i).Ften_deliver_meachul_yyyymm1 + clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm1 + clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm1
    ten_jungsan_yyyymm1 = clsJS.FItemList(i).Ften_jungsan_yyyymm1 + clsJS.FItemList(i+1).Ften_jungsan_yyyymm1 + clsJS.FItemList(i+2).Ften_jungsan_yyyymm1
    ext_jungsan_yyyymm1 = clsJS.FItemList(i).Fext_jungsan_yyyymm1 + clsJS.FItemList(i+1).Fext_jungsan_yyyymm1 + clsJS.FItemList(i+2).Fext_jungsan_yyyymm1

    ten_meachul_yyyymm2 = clsJS.FItemList(i).Ften_meachul_yyyymm2 + clsJS.FItemList(i+1).Ften_meachul_yyyymm2 + clsJS.FItemList(i+2).Ften_meachul_yyyymm2
    ten_deliver_meachul_yyyymm2 = clsJS.FItemList(i).Ften_deliver_meachul_yyyymm2 + clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm2 + clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm2
    ten_jungsan_yyyymm2 = clsJS.FItemList(i).Ften_jungsan_yyyymm2 + clsJS.FItemList(i+1).Ften_jungsan_yyyymm2 + clsJS.FItemList(i+2).Ften_jungsan_yyyymm2
    ext_jungsan_yyyymm2 = clsJS.FItemList(i).Fext_jungsan_yyyymm2 + clsJS.FItemList(i+1).Fext_jungsan_yyyymm2 + clsJS.FItemList(i+2).Fext_jungsan_yyyymm2

    ten_meachul_yyyymm3 = clsJS.FItemList(i).Ften_meachul_yyyymm3 + clsJS.FItemList(i+1).Ften_meachul_yyyymm3 + clsJS.FItemList(i+2).Ften_meachul_yyyymm3
    ten_deliver_meachul_yyyymm3 = clsJS.FItemList(i).Ften_deliver_meachul_yyyymm3 + clsJS.FItemList(i+1).Ften_deliver_meachul_yyyymm3 + clsJS.FItemList(i+2).Ften_deliver_meachul_yyyymm3
    ten_jungsan_yyyymm3 = clsJS.FItemList(i).Ften_jungsan_yyyymm3 + clsJS.FItemList(i+1).Ften_jungsan_yyyymm3 + clsJS.FItemList(i+2).Ften_jungsan_yyyymm3
    ext_jungsan_yyyymm3 = clsJS.FItemList(i).Fext_jungsan_yyyymm3 + clsJS.FItemList(i+1).Fext_jungsan_yyyymm3 + clsJS.FItemList(i+2).Fext_jungsan_yyyymm3

    ten_meachul_null = clsJS.FItemList(i).Ften_meachul_null + clsJS.FItemList(i+1).Ften_meachul_null
    ten_deliver_meachul_null = clsJS.FItemList(i).Ften_deliver_meachul_null + clsJS.FItemList(i+1).Ften_deliver_meachul_null
    ten_jungsan_null = clsJS.FItemList(i).Ften_jungsan_null + clsJS.FItemList(i+1).Ften_jungsan_null
    ext_jungsan_null = clsJS.FItemList(i).Fext_jungsan_null + clsJS.FItemList(i+1).Fext_jungsan_null

    ten_meachul_sum = clsJS.FItemList(i).Ften_meachul_sum + clsJS.FItemList(i+1).Ften_meachul_sum + clsJS.FItemList(i+2).Ften_meachul_sum
    ten_deliver_meachul_sum = clsJS.FItemList(i).Ften_deliver_meachul_sum + clsJS.FItemList(i+1).Ften_deliver_meachul_sum + clsJS.FItemList(i+2).Ften_deliver_meachul_sum
    ten_jungsan_sum = clsJS.FItemList(i).Ften_jungsan_sum + clsJS.FItemList(i+1).Ften_jungsan_sum + clsJS.FItemList(i+2).Ften_jungsan_sum
    ext_jungsan_sum = clsJS.FItemList(i).Fext_jungsan_sum + clsJS.FItemList(i+1).Fext_jungsan_sum + clsJS.FItemList(i+2).Fext_jungsan_sum

    tot_ten_meachul_yyyymm0 = tot_ten_meachul_yyyymm0 + ten_meachul_yyyymm0
    tot_ten_deliver_meachul_yyyymm0 = tot_ten_deliver_meachul_yyyymm0 + ten_deliver_meachul_yyyymm0
    tot_ten_jungsan_yyyymm0 = tot_ten_jungsan_yyyymm0 + ten_jungsan_yyyymm0
    tot_ext_jungsan_yyyymm0 = tot_ext_jungsan_yyyymm0 + ext_jungsan_yyyymm0

    tot_ten_meachul_yyyymm1 = tot_ten_meachul_yyyymm1 + ten_meachul_yyyymm1
    tot_ten_deliver_meachul_yyyymm1 = tot_ten_deliver_meachul_yyyymm1 + ten_deliver_meachul_yyyymm1
    tot_ten_jungsan_yyyymm1 = tot_ten_jungsan_yyyymm1 + ten_jungsan_yyyymm1
    tot_ext_jungsan_yyyymm1 = tot_ext_jungsan_yyyymm1 + ext_jungsan_yyyymm1

    tot_ten_meachul_yyyymm2 = tot_ten_meachul_yyyymm2 + ten_meachul_yyyymm2
    tot_ten_deliver_meachul_yyyymm2 = tot_ten_deliver_meachul_yyyymm2 + ten_deliver_meachul_yyyymm2
    tot_ten_jungsan_yyyymm2 = tot_ten_jungsan_yyyymm2 + ten_jungsan_yyyymm2
    tot_ext_jungsan_yyyymm2 = tot_ext_jungsan_yyyymm2 + ext_jungsan_yyyymm2

    tot_ten_meachul_yyyymm3 = tot_ten_meachul_yyyymm3 + ten_meachul_yyyymm3
    tot_ten_deliver_meachul_yyyymm3 = tot_ten_deliver_meachul_yyyymm3 + ten_deliver_meachul_yyyymm3
    tot_ten_jungsan_yyyymm3 = tot_ten_jungsan_yyyymm3 + ten_jungsan_yyyymm3
    tot_ext_jungsan_yyyymm3 = tot_ext_jungsan_yyyymm3 + ext_jungsan_yyyymm3

    tot_ten_meachul_null = tot_ten_meachul_null + ten_meachul_null
    tot_ten_deliver_meachul_null = tot_ten_deliver_meachul_null + ten_deliver_meachul_null
    tot_ten_jungsan_null = tot_ten_jungsan_null + ten_jungsan_null
    tot_ext_jungsan_null = tot_ext_jungsan_null + ext_jungsan_null

    tot_ten_meachul_sum = tot_ten_meachul_sum + ten_meachul_sum
    tot_ten_deliver_meachul_sum = tot_ten_deliver_meachul_sum + ten_deliver_meachul_sum
    tot_ten_jungsan_sum = tot_ten_jungsan_sum + ten_jungsan_sum
    tot_ext_jungsan_sum = tot_ext_jungsan_sum + ext_jungsan_sum

    %>

	<tr bgcolor="#f6f6f6" align="right">
		<td nowrap bgcolor="#f6f6f6" align="center">합계</td>
		<td width="80"><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S')" style="cursor:pointer;<%if clsJS.FItemList(i).Fyyyymm < stDate and  clsJS.FItemList(i).Fyyyymm > edDate then%>color:gray;<%end if%>"><%= FormatNumber(payPrice, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>

        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",0,mindate)%>','','S', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S', '<%=dateadd("m",0,mindate)%>')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_yyyymm0, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",0,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ext_jungsan_yyyymm0, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",1,mindate)%>','','S', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_yyyymm1, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S', '<%=dateadd("m",1,mindate)%>')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_yyyymm1, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",1,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ext_jungsan_yyyymm1, 0) %></span></td>

        <% if (diffdate >= 2) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",2,mindate)%>','','S', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_yyyymm2, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S', '<%=dateadd("m",2,mindate)%>')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_yyyymm2, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",2,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ext_jungsan_yyyymm2, 0) %></span></td>
        <% end if %>
        <% if (diffdate >= 3) then %>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','<%=dateadd("m",3,mindate)%>','','S', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_yyyymm3, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S', '<%=dateadd("m",3,mindate)%>')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_yyyymm3, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','<%=dateadd("m",3,mindate)%>','S', '')" style="cursor:pointer;"><%= formatnumber(ext_jungsan_yyyymm3, 0) %></span></td>
        <% end if %>

        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','N','','S', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_null, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','S', 'N')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_null, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','N','S', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_null, 0) %></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(ten_meachul_sum, 0) %></span></td>
        <td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(ten_deliver_meachul_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(ten_jungsan_sum, 0) %></span></td>
		<td><span onClick="jsGoDetail('<%= clsJS.FItemList(i).Fsitename %>','<%= clsJS.FItemList(i).Fyyyymm %>','','','I', '')" style="cursor:pointer;"><%= formatnumber(ext_jungsan_sum, 0) %></span></td>
	</tr>
<% end if %>
<% next %>
    <tr bgcolor="#ffffff" align="right">
	 	<td colspan="4" align="center" bgcolor="#E6E6E6">합계</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_yyyymm0,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_yyyymm0,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_yyyymm0,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ext_jungsan_yyyymm0,0)%></b></td>

        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_yyyymm1,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_yyyymm1,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_yyyymm1,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ext_jungsan_yyyymm1,0)%></b></td>

        <% if (diffdate >= 2) then %>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_yyyymm2,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_yyyymm2,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_yyyymm2,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ext_jungsan_yyyymm2,0)%></b></td>
        <% end if %>

        <% if (diffdate >= 2) then %>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_yyyymm3,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_yyyymm3,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_yyyymm3,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ext_jungsan_yyyymm3,0)%></b></td>
        <% end if %>

        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_null,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_null,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_null,0)%></b></td>
        <td width="5" bgcolor="#f4f4f4"></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_meachul_sum,0)%></b></td>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_deliver_meachul_sum,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ten_jungsan_sum,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(tot_ext_jungsan_sum,0)%></b></td>
	 </tr>

</table>
</div>

<p />

[OLD]
<div width="900" style="overflow-x:auto;overflow-y:hidden;">
<table    align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 <Tr bgcolor="#E6E6E6" align="center">
		<td colspan="4">출고/정산일</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate %>
		<td colspan="3"><%=left(dateadd("m",i,mindate),7)%></td>
		<%next%>
		<td colspan="2">미매칭</td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td colspan="3">계</td>
	</tr>
	<Tr bgcolor="#E6E6E6" align="center">
		<td width="50" nowrap>결제일</td>
		<td width="50" nowrap>제휴몰</td>
		<td width="50" nowrap>구분</td>
		<TD width="80" nowrap>결제(승인)액</TD>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate %>
	 	<TD width="80" nowrap>10x10출고액</TD>
		<TD width="80">(10x10)<br/>제휴정산액</TD>
	 	<TD width="80">(제휴)<br/>제휴정산액</TD>
		<%next%>
		<TD width="80" nowrap>10x10출고액</TD>
	 	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	 	<!--<TD width="80" >(제휴)<br/>제휴정산액</TD>-->
		<td width="5" bgcolor="#f4f4f4"></td>

		<TD width="80" nowrap>10x10출고액</TD>
	 	<TD width="80" >(10x10)<br/>제휴정산액</TD>
	 	<TD width="80" >(제휴)<br/>제휴정산액</TD>
	</tr>
	<% dim totscmI, totsomI, totomI
		dim totscmD, totsomD, totomD
		dim totscmM, totsomM, totomM
        dim sumTotS, sumTotOS, sumTotO, sumSN , sumOSN
        dim sumS(), sumOS(), sumO()
		 sumTotS=0: sumTotOS=0: sumTotO=0
         sumSN=0: sumOSN=0
         redim preserve sumS(diffdate)
         redim preserve sumOS(diffdate)
         redim preserve sumO(diffdate)
         for i = 0 To diffdate
          sumS(i)=0
          sumOS(i)=0
          sumO(i)=0
         next
if isArray(arrList) then
	For intLoop = 0 To ubound(arrList,2)
         totscmI=0 :	totsomI=0: totomI =0
		 totscmD=0: totsomD=0: totomD =0
		 totscmM=0: totsomM=0: totomM=0

    %>
	<tr bgcolor="#FFFFFF" align="right">
		<td rowspan="4"  width="50" bgcolor="#f6f6f6" align="center"><%=arrList(1,intLoop)%></td>
		<td rowspan="4"  width="50" bgcolor="#f6f6f6" align="center"><%=arrList(0,intLoop)%></td>
		<td nowrap  width="50" bgcolor="#f6f6f6" align="center">상품</td>
		<td width="80"><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;<%if arrList(1,intLoop) < stDate and  arrList(1,intLoop) > edDate then%>color:gray;<%end if%>"><%=formatnumber(arrList(2,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate
			totscmI = totscmI + arrList(11+(i*9),intLoop)
			totsomI = totsomI + arrList(14+(i*9),intLoop)
			totomI = totomI + arrList(17+(i*9),intLoop)
		%>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','<%=dateadd("m",i,mindate)%>','','I')" style="cursor:pointer;"><%=formatnumber(arrList(11+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','I')" style="cursor:pointer;"><%=formatnumber(arrList(14+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','I')" style="cursor:pointer;"><%=formatnumber(arrList(17+(i*9),intLoop),0)%></span></td>
		<%Next
			totscmI = totscmI + arrList(5,intLoop)
			totsomI = totsomI + arrList(8,intLoop)
		%>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','N','','I')" style="cursor:pointer;"><%=formatnumber(arrList(5,intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','N','I')" style="cursor:pointer;"><%=formatnumber(arrList(8,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totscmI,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totsomI,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totomI,0)%></span></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="right">
		<td nowrap width="50" bgcolor="#f6f6f6" align="center">배송비</td>
		<td width="50"><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','D')" style="cursor:pointer;<%if arrList(1,intLoop) < stDate then%>color:gray;<%end if%>"><%=formatnumber(arrList(3,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate
			totscmD = totscmD + arrList(12+(i*9),intLoop)
			totsomD = totsomD + arrList(15+(i*9),intLoop)
			totomD= totomD + arrList(18+(i*9),intLoop)
		%>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','<%=dateadd("m",i,mindate)%>','','D')" style="cursor:pointer;"><%=formatnumber(arrList(12+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','D')" style="cursor:pointer;"><%=formatnumber(arrList(15+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','D')" style="cursor:pointer;"><%=formatnumber(arrList(18+(i*9),intLoop),0)%></span></td>
		<%Next
			totscmD = totscmD + arrList(6,intLoop)
			totsomD = totsomD + arrList(9,intLoop)
		%>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','N','','D')" style="cursor:pointer;"><%=formatnumber(arrList(6,intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','N','D')" style="cursor:pointer;"><%=formatnumber(arrList(9,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(totscmD,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(totsomD,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','D')" style="cursor:pointer;"><%=formatnumber(totomD,0)%></span></td>
	</tr>
	<tr bgcolor="#FFFFFF" align="right">
		<td nowrap bgcolor="#f6f6f6" align="center">+/-취소액
		<%if arrList(1,intLoop) >= stDate and arrList(1,intLoop)<=edDate   then%>
		<div style="color:gray;">[미매칭취소액]</div>
		<%end if%>
		</td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','M')" style="cursor:pointer;<%if arrList(1,intLoop) < stDate then%>color:gray;<%end if%>"><%=formatnumber(arrList(4,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate
			totscmM = totscmM + arrList(13+(i*9),intLoop)
			totsomM = totsomM + arrList(16+(i*9),intLoop)
			totomM = totomM + arrList(19+(i*9),intLoop)
		%>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','<%=dateadd("m",i,mindate)%>','','M')" style="cursor:pointer;"><%=formatnumber(arrList(13+(i*9),intLoop),0)%></span>
		<%if arrList(7,intLoop) <> 0 and arrList(1,intLoop) = left(dateadd("m",i,mindate),7) then%><div onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','N','','M')" style="cursor:pointer;color:gray">[<%=formatnumber(arrList(7,intLoop),0)%>]</div><%end if%>
		</td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','M')" style="cursor:pointer;"><%=formatnumber(arrList(16+(i*9),intLoop),0)%></span>
		<%if arrList(10,intLoop) <> 0 and arrList(1,intLoop) = left(dateadd("m",i,mindate),7)  then%><div onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','N','M')" style="cursor:pointer;color:gray">[<%=formatnumber(arrList(10,intLoop),0)%>]</div><%end if%>
		</td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','M')" style="cursor:pointer;"><%=formatnumber(arrList(19+(i*9),intLoop),0)%></span></td>
		<%Next
			totscmM = totscmM + arrList(7,intLoop)
			totsomM = totsomM + arrList(10,intLoop)
		%>
		<td></td>
		<td></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','M')" style="cursor:pointer;"><%=formatnumber(totscmM,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','M')" style="cursor:pointer;"><%=formatnumber(totsomM,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','M')" style="cursor:pointer;"><%=formatnumber(totomM,0)%></span></td>
	</tr>
	<tr bgcolor="#f6f6f6" align="right">
		<td nowrap bgcolor="#f6f6f6" align="center">계</td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','S')" style="cursor:pointer;<%if arrList(1,intLoop) < stDate then%>color:gray;<%else%>font-weight:bold;<%end if%>"><%=formatnumber(arrList(2,intLoop)+arrList(3,intLoop)+arrList(4,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<%for i = 0 To diffdate
        sumS(i) = sumS(i) + arrList(11+(i*9),intLoop)+arrList(12+(i*9),intLoop)+arrList(13+(i*9),intLoop)
        sumOS(i) = sumOS(i) + arrList(14+(i*9),intLoop)+arrList(15+(i*9),intLoop)+arrList(16+(i*9),intLoop)
        sumO(i)= sumO(i) + arrList(17+(i*9),intLoop)+arrList(18+(i*9),intLoop)+arrList(19+(i*9),intLoop)
        %>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','<%=dateadd("m",i,mindate)%>','','S')" style="cursor:pointer;"><%=formatnumber(arrList(11+(i*9),intLoop)+arrList(12+(i*9),intLoop)+arrList(13+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','S')" style="cursor:pointer;"><%=formatnumber(arrList(14+(i*9),intLoop)+arrList(15+(i*9),intLoop)+arrList(16+(i*9),intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','<%=dateadd("m",i,mindate)%>','S')" style="cursor:pointer;"><%=formatnumber(arrList(17+(i*9),intLoop)+arrList(18+(i*9),intLoop)+arrList(19+(i*9),intLoop),0)%></span></td>
		<%Next
      '  sumSN = sumSN + arrList(5,intLoop)+arrList(6,intLoop)+arrList(7,intLoop)
	    sumSN = sumSN + arrList(5,intLoop)+arrList(6,intLoop)
      '  sumOSN = sumOSN + arrList(8,intLoop)+arrList(9,intLoop)+arrList(10,intLoop)
	   sumOSN = sumOSN + arrList(8,intLoop)+arrList(9,intLoop)
        %>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','N','','S')" style="cursor:pointer;"><%=formatnumber(arrList(5,intLoop)+arrList(6,intLoop),0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','N','S')" style="cursor:pointer;"><%=formatnumber(arrList(8,intLoop)+arrList(9,intLoop),0)%></span></td>
		<td width="5" bgcolor="#f4f4f4"></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totscmI+totscmD+totscmM,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totsomI+totsomD+totsomM,0)%></span></td>
		<td><span onClick="jsGoDetail('<%=arrList(0,intLoop)%>','<%=arrList(1,intLoop)%>','','','I')" style="cursor:pointer;"><%=formatnumber(totomI+totomD+totomM,0)%></span></td>
	</tr>
	<%Next%>
    <tr bgcolor="#ffffff" align="right">
	 	<td colspan="4" align="center" bgcolor="#E6E6E6">합계</td>
		<td width="5" bgcolor="#f4f4f4"></td>
        <%for i = 0 To diffdate
        sumTotS = sumTotS + sumS(i)
        sumTotOS= sumTotOS + sumOS(i)
        sumTotO = sumTotO + sumO(i)
         %>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumS(i),0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumOS(i),0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumO(i),0)%></b></td>
		<%Next
         sumTotS = sumTotS + sumSN
        sumTotOS= sumTotOS + sumOSN
        %>
        <td bgcolor="#f6f6f6"><b><%=formatnumber(sumSN,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumOSN,0)%></b></td>
        <td width="5" bgcolor="#f4f4f4"></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumTotS,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumTotOS,0)%></b></td>
		<td bgcolor="#f6f6f6"><b><%=formatnumber(sumTotO,0)%></b></td>
	 </tr>
<%else%>
	<tr><td bgcolor="#ffffff"  colspan="11">등록된 내용이 없습니다.</td></tr>
<%end if%>
    </table>

<iframe id="ifrDB" name="ifrDB" src="about:blank" frameborder="0" width="100%" height="300"></iframe>
<!-- 검색 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
