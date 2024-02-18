<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 조짐 상품 
' Hieditor : 서동석 생성
'			 2020.03.17 한용민 수정(검색조건 전시카테고리 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, vArr1, i, j, k, dispCate
Dim vSDate, vEDate, vChannel, oTp, itopn, ptime, makerid, mwdiv, onlysoldout, onlynv
dispCate = requestCheckvar(request("disp"),16)
vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
oTp  = requestCheckvar(request("oTp"),10)
itopn = requestCheckvar(request("itopn"),10)
ptime = requestCheckvar(request("ptime"),10)
makerid = requestCheckvar(request("makerid"),32)
mwdiv = requestCheckvar(request("mwdiv"),10)
onlysoldout = requestCheckvar(request("onlysoldout"),10)
onlynv = requestCheckvar(request("onlynv"),10)


if (oTp="") then oTp="b" ''장바구니(b),  위시순(w), 주문순(o), 조회순(v)
if (itopn="") then itopn=100
if (ptime="") then ptime="-999"
         
If vSDate = "" Then
	vSDate = LEFT(dateadd("d",-0,Date()),10)
End If

If vEDate = "" Then
	vEDate = LEFT(date(),10)
End If

Dim thedate 
if (ptime="-1") then
	thedate=LEFT(dateadd("d",-1,now()),10)
elseif (ptime="-2") then
	thedate=LEFT(dateadd("d",-2,now()),10)
elseif (ptime="-3") then
	thedate=LEFT(dateadd("d",-3,now()),10)
elseif (ptime="-999") then
	thedate=vSDate
end if

Dim iszozimtype : iszozimtype = 1
if ((oTp="ZD") or (oTp="ZU")) then iszozimtype=2	'' 상품
if ((oTp="BD") or (oTp="BU")) then iszozimtype=3  	'' 브랜드

SET oChart = new CChart
	''oChart.FRectSDate = vSDate
	''oChart.FRectEDate = vEDate
	oChart.FPageSize = itopn
	oChart.FRectPreTime = ptime
	oChart.FRectTheDate = thedate
	oChart.FRectOrderType = oTp
	oChart.FRectMakerid = makerid
	oChart.FRectMwdiv = mwdiv
	oChart.FRectOnlySoldout = onlysoldout
	oChart.FRectOnlyNvShop = CHKIIF(onlynv<>"",1,0)
	oChart.FRectDispCate		= dispCate
	if (iszozimtype=1) then
		vArr1 = oChart.fnRequireConversionItem_DW()
	elseif  (iszozimtype=2) then
		vArr1 = oChart.fnZoomUpDownItem_DW()
	elseif  (iszozimtype=3) then
		vArr1 = oChart.fnZoomUpDownBrand_DW()
	end if
SET oChart = nothing


dim imgURL, iSellyn, iLimityn, iLimitNo, isellStr, iLimitStr
%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>


<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type="text/javascript">
$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
		//	CAL_End.args.min = date;
		//	CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	/*
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	*/
});

function goSearch(){
	if($("#sdate").val() == ""){
		alert("시작일을 입력하세요");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("종료일을 입력하세요");	
		return false;
	}
	document.frm1.submit();
}

//상품판매추이그래프
function popItemSellGraph(itemid) {
	var popItemSellGraph = window.open("/admin/maechul/item_graph.asp?itemid="+itemid,"popItemSellGraph","width=1400, height=1000,resizable=yes, scrollbars=yes");
	popItemSellGraph.focus();
}

function popItemTrend(itemid){
	var popwin = window.open("/admin/maechul/itemTrend.asp?itemid="+itemid,"popItemTrend","width=1400, height=1000,resizable=yes, scrollbars=yes");
	popwin.focus();
}

//브랜드 추세
function popBrandSellGraph(makerid,startdate,enddate) {
	var popBrandSellGraph = window.open("/admin/dataanalysis/chart/sellbybrand.asp?pvalue="+makerid+"&startdate="+startdate+"&enddate="+enddate,"popBrandSellGraph","width=1700, height=800,resizable=yes, scrollbars=yes");
	popBrandSellGraph.focus();
}

function chgComp(comp){
	var ival = comp.value;
	if (ival=="-999"){
		document.getElementById("datebox").style.display="";
	}else{
		document.getElementById("datebox").style.display="none";
	}
}

function setEnDisable(comp){
	var ival = comp.value;
	
	var selval = document.frm1.ptime.options[document.frm1.ptime.selectedIndex].value;

	if (((ival=="ZD")||(ival=="ZU")||(ival=="BD")||(ival=="BU"))&&(selval!="-999")){
		document.frm1.ptime.value="-999";
		document.getElementById("datebox").style.display="";

	}

	if ((ival=="ZD")||(ival=="ZU")){
		document.getElementById("idmwbox").style.display="";
	}else{
		document.getElementById("idmwbox").style.display="none";
	}

	if ((ival=="ZD")||(ival=="ZU")||(ival=="BD")||(ival=="BU")){
		document.getElementById("invbox").style.display="";
	}else{
		document.getElementById("invbox").style.display="none";
	}

	
	
}
</script>


<body>
<form name="frm1" method="post" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	
    
    검색 기간 : 
    <% CALL drawPreTimeSelectBox("ptime",ptime) %>
    &nbsp;&nbsp;

	<span id="datebox" name="datebox" style="display:<%=CHKIIF(ptime="-999","","none")%>">
	날짜(주문일) : 
	<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
	<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
	</span>		
	
	
    &nbsp;&nbsp;
    <input type="radio" name="oTp" value="b" <%=CHKIIF(oTp="b","checked","") %> onClick="setEnDisable(this);">장바구니순
    <input type="radio" name="oTp" value="w" <%=CHKIIF(oTp="w","checked","") %> onClick="setEnDisable(this);">위시순
    <input type="radio" name="oTp" value="v" <%=CHKIIF(oTp="v","checked","") %> onClick="setEnDisable(this);">조회순
    <input type="radio" name="oTp" value="o" <%=CHKIIF(oTp="o","checked","") %> onClick="setEnDisable(this);">주문순
	&nbsp;
	|
	&nbsp;
	<input type="radio" name="oTp" value="ZD" <%=CHKIIF(oTp="ZD","checked","") %> onClick="setEnDisable(this);">판매급락 상품
	<input type="radio" name="oTp" value="ZU" <%=CHKIIF(oTp="ZU","checked","") %> onClick="setEnDisable(this);">판매급등 상품

    &nbsp;
	|
	&nbsp;

	<input type="radio" name="oTp" value="BD" <%=CHKIIF(oTp="BD","checked","") %> onClick="setEnDisable(this);">판매급락 브랜드
	<input type="radio" name="oTp" value="BU" <%=CHKIIF(oTp="BU","checked","") %> onClick="setEnDisable(this);">판매급등 브랜드

    &nbsp;
	|
	&nbsp;
	<input type="radio" name="oTp" value="AA" <%=CHKIIF(oTp="AA","checked","") %> onClick="setEnDisable(this);">조짐a(Testing)
	
	&nbsp;&nbsp;
   
    
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>
<tr bgcolor="#F4F4F4">
	<td>
	건수
	<select name="itopn">
		<option value="100" <%=CHKIIF(itopn="100","selected","") %> >100</option>
		<option value="200" <%=CHKIIF(itopn="200","selected","") %> >200</option>
		<option value="300" <%=CHKIIF(itopn="300","selected","") %> >300</option>
    </select>
	&nbsp;
	브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
	&nbsp;
	<span id="idmwbox" name="idmwbox" style="display:<%=CHKIIF(iszozimtype=2,"","none")%>">
		매입구분 : <% Call drawSelectBoxMWU("mwdiv",mwdiv) %>
		&nbsp;
		<input type="checkbox" name="onlysoldout" <%=CHKIIF(onlysoldout="on","checked","")%>>품절상품만 보기
		&nbsp;
	</span>
	<span id="invbox" name="invbox" style="display:<%=CHKIIF(iszozimtype=2 or iszozimtype=3,"","none")%>">
		<input type="checkbox" name="onlynv" <%=CHKIIF(onlynv="on","checked","")%>>rdsite NvShop 만 보기
	</span>
	&nbsp;
	전시카테고리 : <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
	</td>
</tr>

</table>
</form>
<br />
* 약 1시간 지연데이터
* 주문데이터에 반품 교환건은 포함되지 않음 제휴,해외,3pl은 포함되지 않음, 무통장 결제 이전 주문 포함됨
<p>
<% if (iszozimtype=1) then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<% If isArray(vArr1) Then %>
	<tr bgcolor="#DDDDDD" align="center">
		<td>상품코드</td>
		<td width="50">이미지</td>
		<td>브랜드</td>
		<td>상품명</td>
		
		<% if (oTp="o") then %>
		<td>주문<br>건수</td>
		<td>주문<br>랭킹</td>
		<td>장바구니<br>담은수</td>
		<td>장바구니<br>랭킹</td>
		<td>위시<br>담은수</td>
		<td>위시<br>랭킹</td>
		<td>조회<br>유저수</td>
		<td>조회<br>랭킹</td>
		
		<% elseif (oTp="w") then %>
		<td>위시<br>담은수</td>
		<td>위시<br>랭킹</td>
		<td>장바구니<br>담은수</td>
		<td>장바구니<br>랭킹</td>
		<td>조회<br>유저수</td>
		<td>조회<br>랭킹</td>
		<td>주문<br>건수</td>
		<td>주문<br>랭킹</td>
		
		<% elseif (oTp="v") then %>
		<td>조회<br>유저수</td>
		<td>조회<br>랭킹</td>
		<td>위시<br>담은수</td>
		<td>위시<br>랭킹</td>
		<td>장바구니<br>담은수</td>
		<td>장바구니<br>랭킹</td>
		<td>주문<br>건수</td>
		<td>주문<br>랭킹</td>
		
		<% else %>
		<td>장바구니<br>담은수</td>
		<td>장바구니<br>랭킹</td>
		<td>위시<br>담은수</td>
		<td>위시<br>랭킹</td>
		<td>조회<br>유저수</td>
		<td>조회<br>랭킹</td>
		<td>주문<br>건수</td>
		<td>주문<br>랭킹</td>
		<% end if %>
		
		<td>판매<br>여부</td>
		<td>한정<br>여부</td>
		<td>매출<br>추이</td>
		<td>상품<br>추세</td>
	</tr>
	<% For i = 0 To UBound(vArr1,2) %>
	<%
	imgURL = vArr1(21,i)
	if ((Not IsNULL(imgURL)) and (imgURL<>"")) then 
		imgURL = "<img src='"&webImgUrl & "/image/small/" + GetImageSubFolderByItemid(vArr1(0,i)) + "/"  + vArr1(21,i)&"'>"
	else
		imgURL = ""
	end if

	isellStr =""
	iLimitStr=""

	iSellyn = vArr1(11,i)
	iLimityn = vArr1(12,i)
	iLimitNo = vArr1(13,i)-vArr1(14,i)
	if (iLimitNo<1) then iLimitNo=0
		

	if (iSellyn<>"Y") then isellStr="<strong><font color='#FF0000'>품절</font></strong>"
	if (iSellyn="S") then isellStr="<strong><font color='#CC3333'>일시품절</font></strong>"

	if (iLimityn="Y") then iLimitStr="<font color='#3333CC'>한정<br>("&iLimitNo&")</font>"

	%>
	<tr  bgcolor="#FFFFFF" align="center">
		<td align="left"><%=vArr1(0,i)%></td>
		<td align="left"><%=imgURL%></td>
		<td align="left"><%=vArr1(9,i)%></td>
		<td align="left"><%=vArr1(10,i)%></td>
		
		<td><%=vArr1(1,i)%></td>
		<td><%=vArr1(2,i)%></td>
		<td><%=vArr1(3,i)%></td>
		<td><%=vArr1(4,i)%></td>
		<td><%=vArr1(5,i)%></td>
		<td><%=vArr1(6,i)%></td>
		<td><%=vArr1(7,i)%></td>
		<td><%=vArr1(8,i)%></td>
		
		<td><%=isellStr%></td>
		<td><%=iLimitStr%></td>
		<td><img src="/images/icon_search.jpg" onClick="popItemSellGraph('<%=vArr1(0,i)%>');" style="cursor:pointer"></td>
		<td>
			<img src="/images/icon_search.jpg" onClick="popItemTrend('<%=vArr1(0,i)%>');" style="cursor:pointer">
		</td>
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td>
			검색결과가 없습니다.
		</td>
	</tr>
	<% end if %>
	</table>
<% elseif (iszozimtype=2) then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<% If isArray(vArr1) Then %>
	<tr bgcolor="#DDDDDD" align="center">
		<td>상품코드</td>
		<td width="50">이미지</td>
		<td>브랜드</td>
		<td>상품명</td>
		
		
		<td>이전<br>주문건수</td>
		<td>주문건수</td>
		<td>이전<br>판매수량</td>
		<td>판매수량</td>
		<td>이전<br>구매총액</td>
		<td>구매총액</td>
		
		<td>상승하락<br>지수</td>
		<td>매입<br>구분</td>
		<td>판매<br>여부</td>
		<td>한정<br>여부</td>
		<td>매출<br>추이</td>
		<td>상품<br>추세</td>
	</tr>
	<% For i = 0 To UBound(vArr1,2) %>
	<%
	imgURL = vArr1(18,i)
	if ((Not IsNULL(imgURL)) and (imgURL<>"")) then 
		imgURL = "<img src='"&webImgUrl & "/image/small/" + GetImageSubFolderByItemid(vArr1(1,i)) + "/"  + vArr1(18,i)&"'>"
	else
		imgURL = ""
	end if

	isellStr =""
	iLimitStr=""

	iSellyn = vArr1(16,i)
	iLimityn = vArr1(13,i)
	iLimitNo = vArr1(14,i)-vArr1(15,i)
	if (iLimitNo<1) then iLimitNo=0
		

	if (iSellyn<>"Y") then isellStr="<strong><font color='#FF0000'>품절</font></strong>"
	if (iSellyn="S") then isellStr="<strong><font color='#CC3333'>일시품절</font></strong>"

	if (iLimityn="Y") then iLimitStr="<font color='#3333CC'>한정<br>("&iLimitNo&")</font>"

	%>
	<tr  bgcolor="#FFFFFF" align="center">
		<td align="center"><%=vArr1(1,i)%></td>
		<td align="left"><%=imgURL%></td>
		<td align="left"><%=vArr1(11,i)%></td>
		<td align="left"><%=vArr1(10,i)%></td>
		
		<td ><%=FormatNumber(vArr1(2,i),0)%></td>
		<td ><%=FormatNumber(vArr1(6,i),0)%></td>
		<td ><%=FormatNumber(vArr1(3,i),0)%></td>
		<td ><%=FormatNumber(vArr1(7,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(4,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(8,i),0)%></td>
		<td ><%=FormatNumber(vArr1(0,i),2)%></td>
		<td><font color="<%=mwdivColor(vArr1(17,i))%>"><%=mwdivName(vArr1(17,i))%></font></td>
		<td><%=isellStr%></td>
		<td><%=iLimitStr%></td>
		<td><img src="/images/icon_search.jpg" onClick="popItemSellGraph('<%=vArr1(1,i)%>');" style="cursor:pointer"></td>
		<td>
			<img src="/images/icon_search.jpg" onClick="popItemTrend('<%=vArr1(1,i)%>');" style="cursor:pointer">
		</td>
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td>
			검색결과가 없습니다.
		</td>
	</tr>
	<% end if %>
	</table>
<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<% If isArray(vArr1) Then %>
	<tr bgcolor="#DDDDDD" align="center">
		<td>브랜드</td>
		<td width="50">대표<br>이미지</td>
		<td>브랜드명</td>
		
		
		<td>이전<br>주문건수</td>
		<td>주문건수</td>
		<td>이전<br>판매수량</td>
		<td>판매수량</td>
		<td>이전<br>구매총액</td>
		<td>구매총액</td>
		
		<td>상승하락<br>지수</td>
		<td>판매중상품수</td>
		<td>매출추이</td>
	</tr>
	<% For i = 0 To UBound(vArr1,2) %>
	<%
	imgURL = vArr1(13,i)
	if ((Not IsNULL(imgURL)) and (imgURL<>"") and (Not IsNULL(vArr1(12,i)))) then 
		imgURL = "<img src='"&webImgUrl & "/image/small/" + GetImageSubFolderByItemid(vArr1(12,i)) + "/"  + vArr1(13,i)&"'>"
	else
		imgURL = ""
	end if

	%>
	<tr  bgcolor="#FFFFFF" align="center">
		<td align="center"><%=vArr1(1,i)%></td>
		<td align="left"><%=imgURL%></td>
		<td align="left"><%=vArr1(11,i)%></td>
		
		<td ><%=FormatNumber(vArr1(2,i),0)%></td>
		<td ><%=FormatNumber(vArr1(6,i),0)%></td>
		<td ><%=FormatNumber(vArr1(3,i),0)%></td>
		<td ><%=FormatNumber(vArr1(7,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(4,i),0)%></td>
		<td align="right"><%=FormatNumber(vArr1(8,i),0)%></td>
		<td ><%=FormatNumber(vArr1(0,i),2)%></td>
		<td><%=FormatNumber(vArr1(14,i),0)%></td>
		<td><img src="/images/icon_search.jpg" onClick="popBrandSellGraph('<%=vArr1(1,i)%>','<%=dateadd("d",-7,thedate)%>','<%=thedate%>');" style="cursor:pointer"></td>
	</tr>
	<% next %>
	<% else %>
	<tr bgcolor="#FFFFFF">
		<td>
			검색결과가 없습니다.
		</td>
	</tr>
	<% end if %>
	</table>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->