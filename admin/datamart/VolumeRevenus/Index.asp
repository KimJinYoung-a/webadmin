<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/DataMartItemsalecls.asp"-->

<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2,Param
dim fromDate,toDate,cdL,cdM,research
dim rectoldjumun,dategubun, ckMinus
dim catebase
dim ck2ndcate
Dim quarter, inc3pl
'dim ck_joinmall,ck_ipjummall,ck_pointmall

yyyy1   = RequestCheckVar(request("yyyy1"),4)
mm1     = RequestCheckVar(request("mm1"),2)
dd1     = RequestCheckVar(request("dd1"),2)
yyyy2   = RequestCheckVar(request("yyyy2"),4)
mm2     = RequestCheckVar(request("mm2"),2)
dd2     = RequestCheckVar(request("dd2"),2)
quarter     = RequestCheckVar(request("quarter"),1)

rectoldjumun    = RequestCheckVar(request("rectoldjumun"),10)
dategubun       = ""
research        = RequestCheckVar(request("research"),10)
ckMinus         = RequestCheckVar(request("ckMinus"),10)
catebase  = RequestCheckVar(request("catebase"),10)
''ck_joinmall     = RequestCheckVar(request("ck_joinmall"),10)
''ck_ipjummall    = RequestCheckVar(request("ck_ipjummall"),10)
''ck_pointmall    = RequestCheckVar(request("ck_pointmall"),10)

cdL = RequestCheckVar(request("cd1"),10)
cdM = RequestCheckVar(request("cd2"),10)

ck2ndcate = RequestCheckVar(request("ck2ndcate"),10)
inc3pl = request("inc3pl")

if research<>"on" then
    'if ckMinus="" then ckMinus="1"
	'if ck_joinmall="" then ck_joinmall="on"
	'if ck_ipjummall="" then ck_ipjummall="on"
	'if dategubun="" then dategubun="D"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2,Month(now()))
if (dd1="") then dd1 = Format00(2,day(now()))

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Format00(2,Month(now()))
if (dd2="") then dd2 = Format00(2,day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2)

if (catebase="") then catebase="C"

Param = "&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&dategubun="&dategubun&"&ckMinus="&ckMinus&"&catebase="&catebase


dim oReport
set oReport = new CDatamartItemSale
oReport.FRectStartDate = fromDate
oReport.FRectEndDate = toDate
oReport.FRectDateGubun = dategubun
oReport.FRectIncludeMinus = ckMinus
oReport.FRectCD1 = cdL
oReport.FRectCD2 = cdM
oReport.FRectOldJumun = rectoldjumun
oReport.FRectInclude2ndCate = ck2ndcate
oReport.FRectInc3pl = inc3pl  ''2014/02/24 추가

if (catebase="C") then
    oReport.SearchMallSellrePortChannelByCurrentCateVolumeRevenus   ''현재 카테고리 기준
else
    'oReport.SearchMallSellrePortChannelVolumeRevenus           '' 판매시 카테고리 기준
end if

'oreport.SearchMallSellrePortMonthlyChannel
dim i,p1,p2,p3
dim prename, nextname
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,ch5,ch6,ch7,ch8,ch9,ch10,ch11


dim sellcnt, selltotal, buytotal
dim TTLsellcnt, TTLselltotal, TTLbuytotal
dim TTLorgTotal,TTLitemcostCouponNotApplied

Dim TTvolume , TTrevenus , TTallvolume , TTallrevenus
%>
<script language='javascript'>
function subPage(cd1,cd2){
    frm.cd1.value=cd1;
    frm.cd2.value=cd2;
    frm.submit();
}

function showChart(cdl,cdm,cds){
    var params = 'cdl='+cdl+'&cdm='+cdm+'&cds='+cds+'<%= Param %>';
    var popwin = window.open('/admin/datamart/mkt/cateChartView.asp?' + params,'cateChart','width=950,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function chkinsert(){
	 var popwin = window.open('pop_write.asp','popWrite','width=700,height=400,scrollbars=yes,resizable=yes');
     popwin.focus();
}

function chkComp(comp){
    if (comp.value=="S"){
        comp.form.ck2ndcate.checked=false;
        comp.form.ck2ndcate.disabled=true;
    }else{
        comp.form.ck2ndcate.disabled=false;

    }
}

function chgday(v){
	var frm = document.frm;
	switch(v){
		case  "1" :
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "01";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "03";
			frm.dd2.value = "31";
			frm.submit();
			break;
		case  "2" :
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "04";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "06";
			frm.dd2.value = "30";
			frm.submit();
			break;
		case  "3" :
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "07";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "09";
			frm.dd2.value = "30";
			frm.submit();
			break;
		case "4" :
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "10";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "12";
			frm.dd2.value = "31";
			frm.submit();
			break;
		case "5" : //상반기
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "01";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "06";
			frm.dd2.value = "30";
			frm.submit();
			break;
		case "6" : //하반기
			frm.yyyy1.value = <%=yyyy1%>;
			frm.mm1.value = "07";
			frm.dd1.value = "01";
			frm.yyyy2.value = <%=yyyy2%>;
			frm.mm2.value = "12";
			frm.dd2.value = "31";
			frm.submit();
			break;
	}
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="cd1" value="">
	<input type="hidden" name="cd2" value="">
	<tr>
		<td class="a" >
		<!--
		<input type="checkbox" name="rectoldjumun" <% if rectoldjumun="on" then response.write "checked" %> >6개월이전자료&nbsp;&nbsp;

		<input type="radio" name="dategubun" value="D" <% If dategubun<>"M" Then response.write "checked" %>>일별 <input type="radio" name="dategubun" value="M" <% If dategubun="M" Then response.write "checked" %>>월별
		<br>
		-->
		주문일  :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;
		<select name="ckMinus">
		<option value="" >반품포함
		<option value="1" <%= CHKIIF(ckMinus="1","selected","") %> >반품제외
		<option value="2" <%= CHKIIF(ckMinus="2","selected","") %> >반품주문만
		</select>

		&nbsp;
		카테고리 매출 기준:
		<!-- <input type="radio" name="catebase" value="S" <%= CHKIIF(catebase="S","checked","") %> onClick="chkComp(this)">판매시카테고리 -->
		<input type="radio" name="catebase" value="C" <%= CHKIIF(catebase="C","checked","") %> onClick="chkComp(this)">현재(관리)카테고리

		<select name="ck2ndcate">
		<option value="">기본카테고리
		<option value="All" <%= CHKIIF(ck2ndcate="All","selected","") %>>기본+추가카테고리
		<option value="OnlyA" <%= CHKIIF(ck2ndcate="OnlyA","selected","") %>>추가카테고리
		</select>

        <b>* 매출처구분</b>
        	<% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>

		<td class="a" align="right">
			<a href="javascript:frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	<tr>
		<td class="a">분기별 :
		<input type="radio" name="quarter" id="quarter1" value="1" onclick="chgday('1');" <%=chkiif(CStr(quarter) = "1","checked","")%>><label for="quarter1">1/4분기</label>
		<input type="radio" name="quarter" id="quarter2" value="2" onclick="chgday('2');" <%=chkiif(CStr(quarter) = "2","checked","")%>><label for="quarter2">2/4분기</label>
		<input type="radio" name="quarter" id="quarter3" value="3" onclick="chgday('3');" <%=chkiif(CStr(quarter) = "3","checked","")%>><label for="quarter3">3/4분기</label>
		<input type="radio" name="quarter" id="quarter4" value="4" onclick="chgday('4');" <%=chkiif(CStr(quarter) = "4","checked","")%>><label for="quarter4">4/4분기</label>&nbsp;&nbsp;&nbsp;
		<input type="radio" name="quarter" id="quarter5" value="5" onclick="chgday('5');" <%=chkiif(CStr(quarter) = "5","checked","")%>><label for="quarter5">상반기</label>
		<input type="radio" name="quarter" id="quarter6" value="6" onclick="chgday('6');" <%=chkiif(CStr(quarter) = "6","checked","")%>><label for="quarter6">하반기</label>
		</td>
	</tr>
	</form>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="3" class="a" >
<tr>
	<td>* 보너스 쿠폰, 마일리지 제외됨. 1시간 지연 데이터. 배송비 매출은 제외됨. * 중카테고리는 감성채널만 볼 수 있습니다.</td>
	<td align="right">
		<a href="javascript:chkinsert();"><img src="http://testwebadmin.10x10.co.kr/images/icon_new_registration.gif"  border="0"></a>
    </td>
</tr>
</table>


<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#CCCCCC" class="a" >
    <tr align="center">
      <td width="10%" class="a"><font color="#000">카테고리</font></td>
      <td width="10%" ><font color="#FFFFFF">&nbsp;</font></td>
      <td width="8%" class="a"><font color="#000">건수<br>(상품수량)</font></td>
      <% if (NOT C_InspectorUser) then %>
      <td width="8%" class="a"><font color="#000">소비자가</font></td>
      <td width="7%" class="a"><font color="#000">할인금액</font></td>
      <td width="7%" class="a"><font color="#000">판매가<br>(할인가)</font></td>
      <td width="7%" class="a"><font color="#000">상품쿠폰<br>사용액</font></td>
      <% end if %>
      <td width="12%" class="a"><font color="#000"><strong>구매총액</strong><% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></font><br/><strong>거래액 목표</strong></td>
      
      <td width="8%" class="a"><font color="#000">매입액</font></td>
      <td width="12%" class="a"><font color="#000">수익<br/><br/><strong>수익액 목표</strong></font></td>
      <td width="6%" class="a"><font color="#000">수익율</font></td>
      <td width="8%" class="a"><font color="#000">거래총목표<br/><br/>수익액총목표</font></td>
    </tr>
	<% for i=0 to oreport.FResultCount-1 %>
	<%
		p1 = 0
		if oreport.maxt<>0 then
		    p1 = Clng(oreport.FItemList(i).Fselltotal/oreport.maxt*100)
		end if

		sellcnt		=	sellcnt + oreport.FItemList(i).Fsellcnt
		selltotal	=	selltotal + oreport.FItemList(i).Fselltotal
		buytotal	=	buytotal + oreport.FItemList(i).Fbuytotal
	%>
	<tr bgcolor="#FFFFFF">
	<td align="right" width="120">
			<% IF cdL<>""  and cdM<>"" Then %>
				<%= oReport.FItemList(i).FcateName %>
			<% ElseIf oReport.FItemList(i).FcateCode = "110" then %>
				<a href="javascript:subPage('<%=oReport.FItemList(i).FcateCode %>','')"><%= oReport.FItemList(i).FcateName %></a>
			<% Else %>
				<%= oReport.FItemList(i).FcateName %>
			<% End IF %>

			<% if oreport.FItemList(i).FtotVolume<>0 then %>
				<% If CLng( oreport.FItemList(i).Fselltotal > oreport.FItemList(i).FVolume ) Then %>
				<br/><br/><font color="red">거래액 목표 초과</font>
				<% End If %>
			<% End If %>
			<% if oreport.FItemList(i).FtotRevenus<>0 then %>
				<% If CLng( (oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal) > oreport.FItemList(i).FRevenus ) Then %>
				<br/><br/><font color="blue">수익액 목표 초과</font>
				<% End If %>
			<% End If %>
	</td>
	  <td>
			<table border="0" class="a" width='<%= CStr(p1) %>%' >
			  <tr>
			  	<% if trim(oreport.FItemList(i).FcateCode)="" then %>
			  	<td height='20' background='/images/dot030.gif'>
			  	<% else %>
			  	<td background='/images/dot<%= "0"&right(oreport.FItemList(i).FcateCode,2) %>.gif' height='20' >
			  	<% end if %>
			  	<% if oreport.FTotalPrice<>0 then %>
			  	<%= CLng(oreport.FItemList(i).Fselltotal/oreport.FTotalPrice*10000)/100 %>%
			  	<% end if %>
			  	</td>
			  </tr>
			 </table>
			 <% if oreport.FItemList(i).FtotVolume<>0 then %>
			 <table border="0" class="a" width="<%= CLng(oreport.FItemList(i).Fselltotal/oreport.FItemList(i).FVolume*10000)/100 +50 %>">
			  <tr>
					<td bgcolor="red" height='20'  width="">
					<font color="#FFFFFF">거래 : <%= CLng(oreport.FItemList(i).Fselltotal/oreport.FItemList(i).FVolume*10000)/100 %>%</font>
					</td>
			  </tr>
			</table>
			<% End If %>
			<% if oreport.FItemList(i).FtotRevenus<>0 then %>
			<table border="0" class="a" width="<%= CLng((oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal)/oreport.FItemList(i).FRevenus*10000)/100 +50%>">
			  <tr>
					<td bgcolor="blue" height='20'>
					<font color="#FFFFFF">수익 : <%= CLng((oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal)/oreport.FItemList(i).FRevenus*10000)/100 %>%</font>
					</td>
			  </tr>
			</table>
			<% End If %>
	  </td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum-oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum-oreport.FItemList(i).Fselltotal) %></td>
	    <% end if %>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %><br/><br/>목표 : <%= NullOrCurrFormat(oreport.FItemList(i).FVolume)%></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal,0) %><br/><br/>목표 : <%= NullOrCurrFormat(oreport.FItemList(i).FRevenus)%></td>
	  <td align="center">
	  <% if oreport.FItemList(i).Fselltotal<>0 then %>
	  	<%= 100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FtotVolume)%><br/><br/><%= NullOrCurrFormat(oreport.FItemList(i).FtotRevenus)%></td>
	</tr>

	<%
		TTLsellcnt	= TTLsellcnt + sellcnt
		TTLselltotal= TTLselltotal + selltotal
		TTLbuytotal = TTLbuytotal + buytotal
        TTLorgTotal = TTLorgTotal + oreport.FItemList(i).ForgitemcostSum
        TTLitemcostCouponNotApplied = TTLitemcostCouponNotApplied + oreport.FItemList(i).FitemcostCouponNotAppliedSum

		If oreport.FItemList(i).FVolume = "" Or IsNull(oreport.FItemList(i).FVolume) Then
			oreport.FItemList(i).FVolume = 0
		End If
		If oreport.FItemList(i).FRevenus = "" Or IsNull(oreport.FItemList(i).FRevenus) Then
			oreport.FItemList(i).FRevenus = 0
		End If
		If oreport.FItemList(i).FtotVolume = "" Or IsNull(oreport.FItemList(i).FtotVolume) Then
			oreport.FItemList(i).FtotVolume = 0
		End If
		If oreport.FItemList(i).FtotRevenus = "" Or IsNull(oreport.FItemList(i).FtotRevenus) Then
			oreport.FItemList(i).FtotRevenus = 0
		End If

		TTvolume = Cdbl(TTvolume + oreport.FItemList(i).FVolume)
		TTrevenus = Cdbl(TTrevenus + oreport.FItemList(i).FRevenus)
		TTallvolume = Cdbl(TTallvolume + oreport.FItemList(i).FtotVolume)
        TTallrevenus = Cdbl(TTallrevenus + oreport.FItemList(i).FtotRevenus)

		sellcnt = 0
		selltotal = 0
		buytotal = 0
	%>
	<% next %>
	<tr bgcolor="#FFFFFF"><td colspan="12"></td></tr>
	<tr bgcolor="#FFFFFF">
	  <td align="center">Total</td>
	  <td>
			<% if TTvolume<>0 then %>
			 <table border="0" class="a" width="<%= CLng(TTLselltotal/TTvolume*10000)/100 +50 %>">
			  <tr>
					<td bgcolor="red" height='20'  width="">
					<font color="#FFFFFF">거래 : <%= CLng(TTLselltotal/TTvolume*10000)/100 %>%</font>
					</td>
			  </tr>
			</table>
			<% End If %>
			<% if TTrevenus<>0 then %>
			<table border="0" class="a" width="<%= CLng((TTLselltotal-TTLbuytotal)/TTrevenus*10000)/100 +50%>">
			  <tr>
					<td bgcolor="blue" height='20'>
					<font color="#FFFFFF">수익 : <%= CLng((TTLselltotal-TTLbuytotal)/TTrevenus*10000)/100 %>%</font>
					</td>
			  </tr>
			</table>
			<% End If %>
	  </td>
	  <td align="right"><%= FormatNumber(TTLsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal-TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied-TTLselltotal) %></td>
	<% end if %>
	  <td align="right"><%= FormatNumber(TTLselltotal,0) %><br/><br/>목표 : <%=NullOrCurrFormat(TTvolume)%></td>
	  <td align="right"><%= FormatNumber(TTLbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(TTLselltotal-TTLbuytotal,0) %><br/><br/>목표 : <%=NullOrCurrFormat(TTrevenus)%></td>
	  <td align="center">
	  <% if TTLselltotal<>0 then %>
	  	<%= 100-CLng(TTLbuytotal/TTLselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	  <td align="center"><%=NullOrCurrFormat(TTallvolume)%><br/><br/><%=NullOrCurrFormat(TTallrevenus)%></td>
	</tr>

</table>


<%
set oreport = Nothing
%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->