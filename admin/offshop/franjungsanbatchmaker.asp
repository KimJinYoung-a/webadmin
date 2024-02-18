<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 정산
' History : 서동석 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%
dim page,shopid,mwgubun
dim yyyy1,mm1,nowdate
dim onlymijungsan

mwgubun = requestCheckVar(request("mwgubun"),1)
if mwgubun="" then mwgubun="M"

yyyy1 = requestCheckVar(request("yyyy1"),4)
mm1 = requestCheckVar(request("mm1"),2)
onlymijungsan = requestCheckVar(request("onlymijungsan"),10)

if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), Mid(nowdate,6,2)-1,1)
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
end if

dim ooffsell
set ooffsell = new COffShopSellReport
ooffsell.FRectJungsanYYYY = yyyy1
ooffsell.FRectJungsanMM = mm1
if mwgubun="M" then
	ooffsell.GetFranMeaipJungSanAutoList2
elseif mwgubun="C" then
	ooffsell.GetFranWitak2MeaipChulgoJungSanAutoList
elseif mwgubun="W" then
	ooffsell.GetFranWitakJungSanAutoList
end if

dim i, sum1, cnt1, sum2, sum3
dim realmargin
%>
<script language='javascript'>
function NextStep(idx,v){
	var nextstep ="";
	var segumil ="";
	var ipkumil ="";
	var ret0 = false;

	if (v=="0"){
		nextstep = "1";
	}else if(v=="1"){
		nextstep = "2";
	}else if(v=="2"){
		ret0 = calendarOpen2(document.frmsb.segumil);
		if (ret0!=true){ return };
		ret0 = confirm('세금계산서 발행일' + document.frmsb.segumil.value + ' OK?');
		if (ret0!=true){ return };
		nextstep = "3";
	}else if(v=="3"){
		ret0 = calendarOpen2(document.frmsb.ipkumil);
		if (ret0!=true){ return };
		ret0 = confirm('입금일' + document.frmsb.ipkumil.value + ' OK?');
		if (ret0!=true){ return };
		nextstep = "7";
	}else{
		return;
	}

	var ret = confirm('다음 단계로 진행 하시겠습니까?');
	if (ret){
		document.frmsb.idx.value=idx;
		document.frmsb.currstate.value=nextstep;
		document.frmsb.mode.value = "nextstep2";
		document.frmsb.target = "_blank";
		document.frmsb.submit();
	}
}

function PopJungsanStep(iidx,shopid,jungsanid){
	var popwin = window.open("offjungsanstepedit.asp?menupos=298&idx=" + iidx + "&shopid=" + shopid + "&jungsanid=" + jungsanid,"popjungsanstep","width=720, height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}

function PopEditDetail(iidx){
	var popwin = window.open("offjungsandetailedit.asp?menupos=298&idx=" + iidx ,"popjungsandetail","width=800, height=600,scrollbars=1,resizable=yes");
	popwin.focus();
}

<% if (mwgubun="M")  then %>
function FranAutoJungsanMake(yyyymm,makerid){
	var popwin = window.open("popfranjungsanautomaker.asp?menupos=535&yyyymm=" + yyyymm + "&makerid=" + makerid ,"popofranjungsanautomaker","width=800, height=400,scrollbars=1");
	popwin.focus();
}

function PopFranMeaipList(yyyymm,makerid){
	var popwin = window.open("/admin/fran/upchejumunlist.asp?statecd=9&yyyymm=" + yyyymm + "&designer=" + makerid ,"popfranmeaiplist","width=880, height=400,scrollbars=1,resizable=1");
	popwin.focus();
}
function ReStep2(idx){
	var ret = confirm('업체 확인 완료 상태입니다. 재작성 하시겠습니까?');
	var ret = confirm('정말 재작성 하시겠습니까?');
	if (ret){
		document.frmsb.idx.value=idx;
		document.frmsb.submit();
	}
}
function ReStep(idx){
	var ret = confirm('재작성 하시겠습니까?');
	if (ret){
		document.frmsb.idx.value=idx;
		document.frmsb.submit();
	}
}
<% elseif (mwgubun="C") then %>
function FranAutoJungsanMake(yyyymm,makerid){
	var popwin = window.open("popfranjungsanautomaker2.asp?menupos=535&jgubun=c&yyyymm=" + yyyymm + "&makerid=" + makerid ,"popoffjungsanautomaker","width=800, height=400,scrollbars=1");
	popwin.focus();
}

function ReStep(idx){
	var ret = confirm('재작성 하시겠습니까?');
	if (ret){
		document.frmsb.idx.value=idx;
		document.frmsb.submit();
	}
}

function Make000(iyyyymm, ishopid,ibrandid,iofforfran){
	popwin = window.open('popjungsansummaker.asp?yyyymm='+iyyyymm + '&shopid=' + ishopid + '&makerid=' + ibrandid + '&offorfran=' + iofforfran,'popjungsansummaker','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

<% elseif (mwgubun="W") then %>
function FranAutoJungsanMake(yyyymm,makerid){
	var popwin = window.open("popfranjungsanautomaker2.asp?menupos=535&yyyymm=" + yyyymm + "&makerid=" + makerid ,"popoffjungsanautomaker","width=800, height=400,scrollbars=1");
	popwin.focus();
}

function ReStep(idx){
	var ret = confirm('재작성 하시겠습니까?');
	if (ret){
		document.frmsb.idx.value=idx;
		document.frmsb.submit();
	}
}

function Make000(iyyyymm, ishopid,ibrandid,iofforfran){
	popwin = window.open('popjungsansummaker.asp?yyyymm='+iyyyymm + '&shopid=' + ishopid + '&makerid=' + ibrandid + '&offorfran=' + iofforfran,'popjungsansummaker','width=800,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

<% end if %>
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			정산대상월 : <% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
			매입구분 :
			<input type=radio name=mwgubun value="M" <% if mwgubun="M" then response.write "checked" %> >오프라인 매입
			<input type=radio name=mwgubun value="C" <% if mwgubun="C" then response.write "checked" %> >특정 -> 매입 출고
			<input type=radio name=mwgubun value="W" <% if mwgubun="W" then response.write "checked" %> >특정
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<% if (mwgubun="M") or (mwgubun="C") then %>
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="100">샾구분</td>
	<td width="100">정산ID</td>
	<td width="120">기본마진</td>
	<td width="60">매입건</td>
	<td width="80">소비자가</td>
	<td width="80">매입가</td>
	<td width="80">정산액</td>
	<td width="80">마진</td>
	<td width="120">상태</td>
	<td width="40">삭제</td>
	<td width="50">상세</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<% sum1 = sum1 + ooffsell.FItemList(i).FTotalSellcash %>
<% sum2 = sum2 + ooffsell.FItemList(i).FTotalBuyCash %>
<% sum3 = sum3 + ooffsell.FItemList(i).FRealjungsansum %>
<% cnt1 = cnt1 + ooffsell.FItemList(i).Ftotno %>

<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).FShopid %></td>
	<td><a href="javascript:FranAutoJungsanMake('<%= yyyy1 %>-<%= mm1 %>','<%= ooffsell.FItemList(i).Fchargeuser %>');"><%= ooffsell.FItemList(i).Fchargeuser %></a></td>
	<td align=center>
	<% if InStr(ooffsell.FItemList(i).getChargeDivName,"특정")>0 then %>
	<b><%= ooffsell.FItemList(i).getChargeDivName %> (<%= ooffsell.FItemList(i).FDefaultMargin %>%)</b>
	<% else %>
	<%= ooffsell.FItemList(i).getChargeDivName %> (<%= ooffsell.FItemList(i).FDefaultMargin %>%)
	<% end if %>

	</td>
	<td align=center><a href="javascript:PopFranMeaipList('<%= yyyy1 %>-<%= mm1 %>','<%= ooffsell.FItemList(i).Fchargeuser %>');"><%= FormatNumber(ooffsell.FItemList(i).Ftotno,0) %>건</a></td>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FTotalSellcash,0) %></td>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FTotalBuyCash,0) %></td>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FRealjungsansum,0) %></td>
	<td align=center>
	<% if ooffsell.FItemList(i).FTotalSellcash<>0 then %>
	<% realmargin =	100-CLng(ooffsell.FItemList(i).FRealjungsansum/ooffsell.FItemList(i).FTotalSellcash*100*100)/100 %>
	<% if ooffsell.FItemList(i).FDefaultMargin<>realmargin then %>
		<font color="#BB2222"><%= realmargin  %></font>
	<% else %>
		<%= realmargin  %>
	<% end if %>
	<% end if %>
	</td>
	<td align=center><font color="<%= ooffsell.FItemList(i).GetStateColor %>"><%= ooffsell.FItemList(i).GetCurrStateName %></font></td>
	<td align=center>
	<% if (not IsNULL(ooffsell.FItemList(i).FjungsaMasterIdx))  then %>
		<% if (ooffsell.FItemList(i).FjungsaMasterIdx<>"") and (ooffsell.FItemList(i).FjungsaMasterIdx<>"0") then %>
			<% if ooffsell.FItemList(i).Fcurrstate="2" then %>
			<a href="javascript:ReStep2('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>');">x</a>
			<% elseif (ooffsell.FItemList(i).Fcurrstate=" ") or (ooffsell.FItemList(i).Fcurrstate="0") or (ooffsell.FItemList(i).Fcurrstate="1") then %>
			<a href="javascript:ReStep('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>');">X</a>
			<% end if %>
		<% end if %>
	<% end if %>
	</td>
	<td><a href="javascript:PopEditDetail('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>');">보기</a></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>합계</td>
	<td align=center colspan="2"></td>
	<td align="right"><%= FormatNumber(cnt1,0) %></td>
	<td align="right"><%= FormatNumber(sum1,0) %></td>
	<td align="right"><%= FormatNumber(sum2,0) %></td>
	<td align="right"><%= FormatNumber(sum3,0) %></td>
	<td align=center></td>
	<td align=center></td>
	<td align=center></td>
	<td align=center></td>
</tr>
</table>
<% elseif mwgubun="W" then %>
<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="100">정산ID</td>
	<td width="120">기본마진</td>
	<td width="80">샾구분</td>
	<td width="80">매출건</td>
	<td width="80">매출</td>
	<td width="80">정산액</td>
	<td width="80">마진</td>
	<td width="120">상태</td>
	<td width="40">삭제</td>
	<td width="50">진행</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<%
if ooffsell.FItemList(i).Fjungsantotsum<>0 then
	sum1 = sum1 + ooffsell.FItemList(i).Fjungsantotsum
else
	sum1 = sum1 + ooffsell.FItemList(i).FTotalSellcash
end if
%>

<% sum3 = sum3 + ooffsell.FItemList(i).FRealjungsansum %>
<% cnt1 = cnt1 + ooffsell.FItemList(i).Ftotno %>

<tr bgcolor="#FFFFFF">
	<% if ooffsell.FItemList(i).Fcurrstate="0" or ooffsell.FItemList(i).Fcurrstate="9" then %>
	<td><a href="javascript:Make000('<%= yyyy1 %>-<%= mm1 %>','<%= ooffsell.FItemList(i).FShopID %>','<%= ooffsell.FItemList(i).Fchargeuser %>','FRN');"><%= ooffsell.FItemList(i).Fchargeuser %></a></td>
	<% else %>
	<td><a href="javascript:FranAutoJungsanMake('<%= yyyy1 %>-<%= mm1 %>','<%= ooffsell.FItemList(i).Fchargeuser %>');"><%= ooffsell.FItemList(i).Fchargeuser %></a></td>
	<% end if %>

	<td align=center><%= ooffsell.FItemList(i).getChargeDivName %> (<%= ooffsell.FItemList(i).FDefaultMargin %>%)</td>
	<td align=center><%= ooffsell.FItemList(i).FShopID %></td>
	<% if ooffsell.FItemList(i).Fjungsantotitemcnt<>0 then %>
	<td align=center><%= FormatNumber(ooffsell.FItemList(i).Fjungsantotitemcnt,0) %>건</a></td>
	<% else %>
	<td align=center><%= FormatNumber(ooffsell.FItemList(i).Ftotno,0) %>건</a></td>
	<% end if %>
	<% if ooffsell.FItemList(i).Fjungsantotsum<>0 then %>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).Fjungsantotsum,0) %></td>
	<% else %>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FTotalSellcash,0) %></td>
	<% end if %>
	<td align=right><%= FormatNumber(ooffsell.FItemList(i).FRealjungsansum,0) %></td>
	<td align=center>
	<% if ooffsell.FItemList(i).Fjungsantotsum<>0 then %>
		<% realmargin =	100-CLng(ooffsell.FItemList(i).FRealjungsansum/ooffsell.FItemList(i).Fjungsantotsum*100*100)/100 %>
		<% if ooffsell.FItemList(i).FDefaultMargin<>realmargin then %>
			<font color="#BB2222"><%= realmargin  %></font>
		<% else %>
			<%= realmargin  %>
		<% end if %>
	<% else %>
		<% if ooffsell.FItemList(i).FTotalSellcash<>0 then %>
		<% realmargin =	100-CLng(ooffsell.FItemList(i).FRealjungsansum/ooffsell.FItemList(i).FTotalSellcash*100*100)/100 %>
		<% if ooffsell.FItemList(i).FDefaultMargin<>realmargin then %>
			<font color="#BB2222"><%= realmargin  %></font>
		<% else %>
			<%= realmargin  %>
		<% end if %>
		<% end if %>
	<% end if %>
	</td>
	<td align=center><font color="<%= ooffsell.FItemList(i).GetStateColor %>"><%= ooffsell.FItemList(i).GetCurrStateName %></font></td>
	<td align=center>
	<% if (not IsNULL(ooffsell.FItemList(i).FjungsaMasterIdx)) and (ooffsell.FItemList(i).FjungsaMasterIdx<>"")  then %>
		<% if ooffsell.FItemList(i).Fcurrstate="2" then %>
		<a href="javascript:ReStep2('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>');">x</a>
		<% else %>
		<a href="javascript:ReStep('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>');">X</a>
		<% end if %>
	<% end if %>
	</td>
	<td><a href="javascript:PopJungsanStep('<%= ooffsell.FItemList(i).FjungsaMasterIdx %>','<%= ooffsell.FItemList(i).Fshopid %>','<%= ooffsell.FItemList(i).Fchargeuser %>');">수정</a></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td>합계</td>
	<td align=center></td>
	<td align=center></td>
	<td align="right"><%= FormatNumber(cnt1,0) %></td>
	<td align="right"><%= FormatNumber(sum1,0) %></td>
	<td align="right"><%= FormatNumber(sum3,0) %></td>
	<td align=center></td>
	<td align=center></td>
	<td align=center></td>
	<td align=center></td>
</tr>
</table>
<% end if %>
<form name=frmsb method=post action="dojungsan.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="delmaster">
<input type="hidden" name="segumil" value="">
<input type="hidden" name="ipkumil" value="">
<input type="hidden" name="currstate" value="">

</form>
<%
set ooffsell = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
