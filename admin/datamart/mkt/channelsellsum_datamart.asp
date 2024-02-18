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
dim fromDate,toDate,cdL,cdM,cdS,research,cd4
dim rectoldjumun,dategubun, ckMinus
dim catebase
dim ck2ndcate, inc3pl
'dim ck_joinmall,ck_ipjummall,ck_pointmall

yyyy1   = RequestCheckVar(request("yyyy1"),4)
mm1     = RequestCheckVar(request("mm1"),2)
dd1     = RequestCheckVar(request("dd1"),2)
yyyy2   = RequestCheckVar(request("yyyy2"),4)
mm2     = RequestCheckVar(request("mm2"),2)
dd2     = RequestCheckVar(request("dd2"),2)

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
cdS = RequestCheckVar(request("cd3"),10)
cd4 = RequestCheckVar(request("cd4"),10)

ck2ndcate = RequestCheckVar(request("ck2ndcate"),10)
inc3pl = request("inc3pl")

if research<>"on" then
    ''if ckMinus="" then ckMinus="1"
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
toDate = DateSerial(yyyy2, mm2, dd2+1)

if (catebase="") then catebase="V"

Param = "&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&dategubun="&dategubun&"&ckMinus="&ckMinus&"&catebase="&catebase


dim oReport
set oReport = new CDatamartItemSale
oReport.FRectStartDate = fromDate
oReport.FRectEndDate = toDate
oReport.FRectDateGubun = dategubun
oReport.FRectIncludeMinus = ckMinus
oReport.FRectCD1 = cdL
oReport.FRectCD2 = cdM
oReport.FRectCD3 = cdS
oReport.FRectCD4 = cd4
oReport.FRectOldJumun = rectoldjumun
oReport.FRectInclude2ndCate = ck2ndcate
oReport.FRectInc3pl = inc3pl  ''2014/02/24 추가

if (catebase="V") then
    oReport.SearchMallSellrePortChannelByCurrentDispCate    ''신 전시카테고리 기준
elseif (catebase="C") then
    oReport.SearchMallSellrePortChannelByCurrentCate   ''현재(관리) 카테고리 기준
else
    oReport.SearchMallSellrePortChannel           '' 판매시(관리) 카테고리 기준
end if

'oreport.SearchMallSellrePortMonthlyChannel
dim i,p1,p2
dim prename, nextname
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,ch5,ch6,ch7,ch8,ch9,ch10,ch11


dim sellcnt, selltotal, buytotal
dim TTLsellcnt, TTLselltotal, TTLbuytotal
dim TTLorgTotal,TTLitemcostCouponNotApplied
%>
<script language='javascript'>
function subPage(cd1,cd2,cd3,cd4){
    var frm=document.frm;

    frm.cd1.value=cd1;
    frm.cd2.value=cd2;
    frm.cd3.value=cd3;
    frm.cd4.value=cd4;
    frm.submit();
}

function subBest(cd1,cd2,cd3,cd4){
    var frm1 = document.frmBuf;

    frm1.cd1.value=cd1;
    frm1.cd2.value=cd2;
    frm1.cd3.value=cd3;
    frm1.cd4.value=cd4;
    frm1.action="channelsellsum_datamart_best.asp";
    frm1.target="channelBest";
    frm1.submit();
}

function showChart(cdl,cdm,cds,cd4){


    <% IF(catebase="V") then%>
    var params = 'disp='+cdl+cdm+cds+cd4+'<%= Param %>';
    var popwin = window.open('cateChartView_DispCate.asp?' + params,'cateChart','width=950,height=800,scrollbars=yes,resizable=yes');
    <% else %>
    var params = 'cdl='+cdl+'&cdm='+cdm+'&cds='+cds+'<%= Param %>';
    var popwin = window.open('cateChartView.asp?' + params,'cateChart','width=950,height=800,scrollbars=yes,resizable=yes');
    <% end if %>
    popwin.focus();
}


//그리드 관련.
function initGrid(){
    var Grid1 = document.all.TnGrid;

    if (!Grid1) return;

    //Grid1.setSelectMode(isRowSelect); // 디폴트 0 : Col select , set 1 id Rowselect
    Grid1.setSelectMode(1);
    Grid1.setDefaultRowHeight(24);

    //Grid1.addNewColumn(KyeName,Caption,Width,Type,Editable,Alignment); TAlignment = (taLeftJustify, taRightJustify, taCenter)
    //TEditorType = ('text','');

    Grid1.addNewColumn('CATEGUBUN','카테고리',140,'text',0,2);
    Grid1.addNewColumn('GRAPH','그래프',300,'text',0,2);
    Grid1.addNewColumn('ORDERCNT','건수',90,'text',0,2);
    Grid1.addNewColumn('SELLSUM','매출액',100,'text',0,1);
    Grid1.addNewColumn('BUYSUM','매입액',100,'text',0,1);
    Grid1.addNewColumn('GAINSUM','수익',100,'text',0,1);
    Grid1.addNewColumn('GAINPRO','수익율',80,'text',0,2);
    Grid1.addNewColumn('BIGO','추세',80,'text',0,1);


    //Grid1.setHiddenColumn('HIDDENTEST');

    //Grid1.setColumnValueColor('VALIDGUBUN','삭제','#FF0000');
    //Grid1.setColumnValueColor('VALIDGUBUN','취소','#FF0000');

    //Grid1.setColumnValueFontStyle('JUMUNDIV','해외','BOLD');

	//Grid1.setColumnValueColor('IPKUMDIVNAME','주문대기','#FF0000');		//0
	//Grid1.setColumnValueColor('IPKUMDIVNAME','주문실패','#44BBBB');		//1
    //Grid1.setColumnValueColor('IPKUMDIVNAME','결제완료','#0000FF');     //4
    //Grid1.setColumnValueColor('IPKUMDIVNAME','주문통보','#CC9933');     //5
    //Grid1.setColumnValueColor('IPKUMDIVNAME','상품준비','#FF00FF');     //6
    //Grid1.setColumnValueColor('IPKUMDIVNAME','일부출고','#EE2222');     //7
    //Grid1.setColumnValueColor('IPKUMDIVNAME','상품출고','#EE2222');     //8
    //Grid1.setColumnValueColor('IPKUMDIVNAME','마이너스','#FF0000');     //9

    //Grid1.setColumnValueRowColor('VALIDGUBUN','삭제','#AAAAAA');
    //Grid1.setColumnValueRowColor('VALIDGUBUN','취소','#AAAAAA');


}

//window.onload = initGrid;

function gridQuery(page){
    ///lib/util/gridResponse.asp?cmd=channelSellsum&stdt=2009-01-01&eddt=2009-06-16
    var Grid1 = document.all.TnGrid;
    var frm   = document.frm;
    var pagesize = 20;

    Grid1.setQueryUrl("<%= manageUrl %>/lib/util/gridResponse.asp");
    Grid1.clearParams();
	Grid1.addParam("cmd", "channelSellsum");
	Grid1.addParam("page", page);
	Grid1.addParam("pagesize", pagesize);

    Grid1.addParam("stdt", frm.yyyy1.value+"-"+frm.mm1.value+"-"+frm.dd1.value);
    Grid1.addParam("eddt", frm.yyyy2.value+"-"+frm.mm2.value+"-"+frm.dd2.value);

	Grid1.addParam("ckMinus",frm.ckMinus.value);

	//ckMinus
	//Grid1.getWebData(parsingType);
	try {
	    Grid1.getWebData(1);
	}catch(e){
	    alert(e.description);
	}

}

function chkComp(comp){
    if (comp.value=="C"){
        comp.form.ck2ndcate.disabled=false;
    }else{

        comp.form.ck2ndcate.checked=false;
        comp.form.ck2ndcate.disabled=true;
    }
}

function SubmitForm(frm) {
	if ((CheckDateValid(frm.yyyy1.value, frm.mm1.value, frm.dd1.value) == true) && (CheckDateValid(frm.yyyy2.value, frm.mm2.value, frm.dd2.value) == true)) {
		frm.submit();
	}
}

</script>
<!-- script language="JavaScript" src="/js/tnGrid.js"></script -->

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="cd1" value="">
	<input type="hidden" name="cd2" value="">
	<input type="hidden" name="cd3" value="">
	<input type="hidden" name="cd4" value="">
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
		<input type="radio" name="catebase" value="V" <%= CHKIIF(catebase="V","checked","") %> onClick="chkComp(this)">현재(<strong>전시</strong>)카테고리
		<input type="radio" name="catebase" value="S" <%= CHKIIF(catebase="S","checked","") %> onClick="chkComp(this)">판매시(관리)카테고리
		<input type="radio" name="catebase" value="C" <%= CHKIIF(catebase="C","checked","") %> onClick="chkComp(this)">현재(관리)카테고리

		<select name="ck2ndcate">
		<option value="">기본카테고리
		<option value="All" <%= CHKIIF(ck2ndcate="All","selected","") %>>기본+추가카테고리
		<option value="OnlyA" <%= CHKIIF(ck2ndcate="OnlyA","selected","") %>>추가카테고리
		</select>
		<!--
		<input type="checkbox" name="ck2ndcate" <%= CHKIIF(ck2ndcate="on","checked","") %> <%= CHKIIF(catebase="S","disabled='true'","") %>>추가 카테고리 포함
		-->

		<b>* 매출처구분</b>
        	<% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>

		<td class="a" align="right">
			<a href="javascript:SubmitForm(document.frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<% if (NOT C_InspectorUser) then %>
<table width="100%" border="0" cellspacing="1" cellpadding="3" class="a" >
<tr>
	<td>* 보너스 쿠폰, 마일리지 제외됨. 1시간 지연 데이터. 배송비 매출은 제외됨.</td>
</tr>
</table>
<% end if %>

<!--
<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF" class="a" >
<tr align="center">
    <td>
    <script language='javascript'>
    DrawTnGridTag("TnGrid","100%",400);
    </script>
    </td>
</tr>
</table>
-->


<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#EFBE00" class="a" >
    <tr align="center">
      <td width="160" class="a"><font color="#FFFFFF">카테고리</font></td>
      <td ><font color="#FFFFFF">&nbsp;</font></td>
      <td width="80" class="a"><font color="#FFFFFF">건수<br>(상품수)</font></td>
      <% if (NOT C_InspectorUser) then %>
      <td width="90" class="a"><font color="#FFFFFF">소비자가</font></td>
      <td width="90" class="a"><font color="#FFFFFF">할인금액</font></td>
      <td width="90" class="a"><font color="#FFFFFF">판매가<br>(할인가)</font></td>
      <td width="90" class="a"><font color="#FFFFFF">상품쿠폰<br>사용액</font></td>
      <% end if %>
      <td width="90" class="a"><font color="#FFFFFF"><strong>구매총액</strong><% if (NOT C_InspectorUser) then %><br>(상품쿠폰적용)<% end if %></font></td>
      <td width="90" class="a"><font color="#FFFFFF">매입액</font></td>
      <td width="90" class="a"><font color="#FFFFFF">수익</font></td>
      <td width="80" class="a"><font color="#FFFFFF">수익율</font></td>
      <td width="50" class="a"><font color="#FFFFFF">추세</font></td>
      <td width="50" class="a"><font color="#FFFFFF">상세</font></td>
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
	<td align="left">
	    &nbsp;
			<% IF cdL<>"" and cdM<>"" and cdS<>"" Then %>
				<%= oReport.FItemList(i).FcateName %>
		    <% ElseIF cdL<>"" and cdM<>"" Then %>
				<a href="javascript:subPage('<%= cdL %>','<%= cdM %>','<%=oReport.FItemList(i).FcateCode %>','')"><%= oReport.FItemList(i).FcateName %></a>
			<% ElseIF cdL<>"" Then %>
				<a href="javascript:subPage('<%= cdL %>','<%=oReport.FItemList(i).FcateCode %>','','')"><%= oReport.FItemList(i).FcateName %></a>
			<% Else %>
				<a href="javascript:subPage('<%=oReport.FItemList(i).FcateCode %>','','','')"><%= oReport.FItemList(i).FcateName %></a>
			<% End IF %>
	</td>
	  <td >
			<table border="0" class="a" width='<%= CStr(p1) %>%' >
			  <tr>
			  	<% if trim(oreport.FItemList(i).FcateCode)="" then %>
			  	<td height='20' background='/images/dot030.gif'>
			  	<% else %>
			  	<td background='/images/dot<%= right(oreport.FItemList(i).FcateCode,3) %>.gif' height='20' >
			  	<% end if %>
			  	<% if oreport.FTotalPrice<>0 then %>
			  	<%= CLng(oreport.FItemList(i).Fselltotal/oreport.FTotalPrice*10000)/100 %>%
			  	<% end if %>
			  	</td>
			  </tr>
			</table>
	  </td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).ForgitemcostSum-oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum) %></td>
	  <td align="right"><%= NullOrCurrFormat(oreport.FItemList(i).FitemcostCouponNotAppliedSum-oreport.FItemList(i).Fselltotal) %></td>
	  <% end if %>
	  <td align="right" bgcolor="#7CCE76"><%= FormatNumber(oreport.FItemList(i).Fselltotal,0) %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(oreport.FItemList(i).Fselltotal-oreport.FItemList(i).Fbuytotal,0) %></td>
	  <td align="center">
	  <% if oreport.FItemList(i).Fselltotal<>0 then %>
	  	<%= 100-CLng(oreport.FItemList(i).Fbuytotal/oreport.FItemList(i).Fselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	  
	<% IF cdL<>"" and cdM<>"" and cdS<>"" Then %>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="showChart('<%= cdL %>','<%= cdM %>','<%= cdS %>','<%=oReport.FItemList(i).FcateCode %>');"></td>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="subBest('<%= cdL %>','<%= cdM %>','<%=cdS %>','<%=oReport.FItemList(i).FcateCode %>')"></td>
	<% ElseIF cdL<>"" and cdM<>"" Then %>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="showChart('<%= cdL %>','<%= cdM %>','<%= oReport.FItemList(i).FcateCode %>','');"></td>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="subBest('<%= cdL %>','<%= cdM %>','<%=oReport.FItemList(i).FcateCode %>','')"></td>
	<% ElseIF cdL<>"" Then %>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="showChart('<%= cdL %>','<%= oReport.FItemList(i).FcateCode %>','','');"></td>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="subBest('<%= cdL %>','<%= oReport.FItemList(i).FcateCode %>','','')"></td>
	<% Else %>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="showChart('<%= oReport.FItemList(i).FcateCode %>','','','');"></td>
		<td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="subBest('<%= oReport.FItemList(i).FcateCode %>','','','')"></td>
	<% End IF %>
	  
	</tr>


	<%
		TTLsellcnt	= TTLsellcnt + sellcnt
		TTLselltotal= TTLselltotal + selltotal
		TTLbuytotal = TTLbuytotal + buytotal
        TTLorgTotal = TTLorgTotal + oreport.FItemList(i).ForgitemcostSum
        TTLitemcostCouponNotApplied = TTLitemcostCouponNotApplied + oreport.FItemList(i).FitemcostCouponNotAppliedSum

		sellcnt = 0
		selltotal = 0
		buytotal = 0
	%>

	<% next %>
	<tr bgcolor="#FFFFFF"><td colspan="13"></td></tr>
	<tr bgcolor="#FFFFFF">
	  <td align="center">Total</td>
	  <td></td>
	  <td align="right"><%= FormatNumber(TTLsellcnt,0) %></td>
	  <% if (NOT C_InspectorUser) then %>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLorgTotal-TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied) %></td>
	  <td align="right"><%= NullOrCurrFormat(TTLitemcostCouponNotApplied-TTLselltotal) %></td>
	  <% end if %>
	  <td align="right" bgcolor="#7CCE76"><%= FormatNumber(TTLselltotal,0) %></td>
	  <td align="right"><%= FormatNumber(TTLbuytotal,0) %></td>
	  <td align="right"><%= FormatNumber(TTLselltotal-TTLbuytotal,0) %></td>
	  <td align="center">
	  <% if TTLselltotal<>0 then %>
	  	<%= 100-CLng(TTLbuytotal/TTLselltotal*100*100)/100 %>%
	  <% end if %>
	  </td>
	  <td align="center"><img src="/images/icon_search.jpg" style="cursor:pointer" onClick="showChart('<%= cdL %>','<%= cdM %>','<%= cdS %>');"></td>
	  <td align="center"></td>
	</tr>

</table>

<form name="frmBuf" method="get" action="">
<input type="hidden" name="yyyy1" value="<%=yyyy1%>">
<input type="hidden" name="mm1" value="<%=mm1%>">
<input type="hidden" name="dd1" value="<%=dd1%>">
<input type="hidden" name="yyyy2" value="<%=yyyy2%>">
<input type="hidden" name="mm2" value="<%=mm2%>">
<input type="hidden" name="dd2" value="<%=dd2%>">
<input type="hidden" name="dd2" value="<%=dd2%>">
<input type="hidden" name="cd1" value="<%=cdL%>">
<input type="hidden" name="cd2" value="<%=cdM%>">
<input type="hidden" name="cd3" value="<%=cdS%>">
<input type="hidden" name="cd4" value="<%=cd4%>">
<input type="hidden" name="catebase" value="<%=catebase%>">
<input type="hidden" name="ck2ndcate" value="<%=ck2ndcate%>">
<input type="hidden" name="ckMinus" value="<%=ckMinus%>">
</form>
<%
set oreport = Nothing
%>


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
