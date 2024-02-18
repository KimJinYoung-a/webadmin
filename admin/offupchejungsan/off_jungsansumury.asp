<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 정산 서머리
' History : 2010.05.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
dim yyyy1,mm1, yyyy2, mm2, chkdate, chknotfinish , research, page
dim p_yyyymm, subtotalFlag , i
dim sub_TW_price, sub_UW_price, sub_CM_price, sub_OM_price, sub_SM_price, sub_ET_price
dim ttl_TW_price, ttl_UW_price, ttl_CM_price, ttl_OM_price, ttl_SM_price, ttl_ET_price
dim sub_jungsanprice, sub_waitsum, sub_fixedthissum, sub_fixednextsum, sub_ipkumsum
dim ttl_jungsanprice, ttl_waitsum, ttl_fixedthissum, ttl_fixednextsum, ttl_ipkumsum
	yyyy1       = request("yyyy1")
	mm1         = request("mm1")
	yyyy2       = request("yyyy2")
	mm2         = request("mm2")
	chkdate     = request("chkdate")
	chknotfinish= request("chknotfinish")
	research    = request("research")
	page        = request("page")

if (research="") and (chkdate="") then chkdate="on"
if (page="") then page=1

dim stdt, eddt, StartYYYYMM, EndYYYYMM
if (yyyy1="") then
	stdt = dateserial(year(Now),month(now)-6,1)
	yyyy1 = Left(CStr(stdt),4)
	mm1 = Mid(CStr(stdt),6,2)
	
	eddt = dateadd("d",dateserial(year(Now),month(now)+1,1),-1)
	yyyy2 = Left(CStr(eddt),4)
	mm2 = Mid(CStr(eddt),6,2)
end if

StartYYYYMM = yyyy1 + "-" + mm1
EndYYYYMM   = yyyy2 + "-" + mm2


dim ooffjungsan
set ooffjungsan = new COffJungsan
	ooffjungsan.FRectFixStateExiste = chknotfinish
	
	if (chkdate="on") then
	    ooffjungsan.FRectStartYYYYMM = StartYYYYMM
	    ooffjungsan.FRectEndYYYYMM   = EndYYYYMM
	end if

	ooffjungsan.GetOffJungsanSummary

sub_TW_price = 0
sub_UW_price = 0
sub_CM_price = 0
sub_OM_price = 0
sub_SM_price = 0
sub_ET_price = 0
sub_jungsanprice    = 0
sub_waitsum         = 0
sub_fixedthissum    = 0
sub_fixednextsum    = 0
sub_ipkumsum        = 0
ttl_TW_price = 0
ttl_UW_price = 0
ttl_CM_price = 0
ttl_OM_price = 0
ttl_SM_price = 0
ttl_ET_price = 0
ttl_jungsanprice  = 0
ttl_waitsum         = 0
ttl_fixedthissum    = 0
ttl_fixednextsum    = 0
ttl_ipkumsum        = 0
%>

<script language='javascript'>

function CheckEnabled(frm,comp){
    if (comp.name=='chkdate'){
        frm.chknotfinish.checked=false;
    }else{
        frm.chkdate.checked=false;
    }
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">        
    	<input type="checkbox" name="chkdate" <% if chkdate="on" then response.write "checked" %> onClick="CheckEnabled(frm,this);">
    	&nbsp;기간검색 : <% DrawYMYMBox yyyy1,mm1, yyyy2,mm2 %> (정산 대상월)
    	&nbsp;&nbsp;
		<input type="checkbox" name="chknotfinish" <% if chknotfinish="on" then response.write "checked" %> onClick="CheckEnabled(frm,this);">&nbsp;처리완료 정산월 제외
	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
	       정상발행 : 세금계산서 발행월이 정산월과 같거나 없을때(원천징수)<br>
	       이월발행 : 세금계산서 발행월이 정산월 이후 일때		
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="left">
		총건수:&nbsp;<%= FormatNumber(ooffjungsan.FResultCount,0) %>
	</td>
</tr>
<tr bgcolor="#EEEEEE" align="center">
	<td rowspan="2" width="70">대상월</td>
	<td rowspan="2" width="50">정산일</td>
	<td colspan="6">계약별 구분</td>
	<td rowspan="2" width="100">정산총액</td>
	<td colspan="4">입금진행현황</td>
</tr>
<tr bgcolor="#EEEEEE" align="center">	
	<td width="70">위탁<br>판매</td>
    <td width="70">업체<br>위탁</td>
    <td width="70">오프<br>매입</td>
    <td width="70">매장<br>매입</td>
    <td width="70">출고<br>매입</td>
    <td width="70">기타<br>내역</td>
	<td width="80">확정이전금액</td>
	<td width="80">확정금액<br>(금월결제)</td>
	<td width="80">확정금액<br>(이월결제)</td>
	<td width="80">입금완료금액</td>	
</tr>
<% if ooffjungsan.FresultCount<1 then %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center"  >[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% 
for i=0 to ooffjungsan.FResultCount-1

p_yyyymm = ooffjungsan.FItemList(i).FYYYYMM

sub_TW_price    = sub_TW_price      + ooffjungsan.FItemList(i).FTW_price
sub_UW_price    = sub_UW_price      + ooffjungsan.FItemList(i).FUW_price
sub_CM_price    = sub_CM_price      + ooffjungsan.FItemList(i).FCM_price
sub_OM_price    = sub_OM_price      + ooffjungsan.FItemList(i).FOM_price
sub_SM_price    = sub_SM_price      + ooffjungsan.FItemList(i).FSM_price
sub_ET_price    = sub_ET_price      + ooffjungsan.FItemList(i).FET_price
sub_jungsanprice= sub_jungsanprice  + ooffjungsan.FItemList(i).Ftot_jungsanprice
sub_waitsum     = sub_waitsum       + ooffjungsan.FItemList(i).Fwaitsum
sub_fixedthissum= sub_fixedthissum  + ooffjungsan.FItemList(i).Ffixedthissum
sub_fixednextsum= sub_fixednextsum  + ooffjungsan.FItemList(i).Ffixednextsum
sub_ipkumsum    = sub_ipkumsum      + ooffjungsan.FItemList(i).Fipkumsum

subtotalFlag=false

if (i=ooffjungsan.FResultCount-1) then
    subtotalFlag=true
elseif (ooffjungsan.FItemList(i+1).FYYYYMM<>p_yyyymm) then
    subtotalFlag=true
else
    subtotalFlag=false
end if
%>
<tr bgcolor="#FFFFFF" height=24>
    <td align="center"><%= ooffjungsan.FItemList(i).FYYYYMM %></td>
    <td align="center"><%= ooffjungsan.FItemList(i).Fjungsan_date_off %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FTW_price,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FUW_price,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FOM_price,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FSM_price,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FCM_price,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).FET_price,0) %></td>
    <td align="right">
    <%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %>
    
    <% if (ooffjungsan.FItemList(i).FTW_price + ooffjungsan.FItemList(i).FUW_price + ooffjungsan.FItemList(i).FCM_price + ooffjungsan.FItemList(i).FOM_price + ooffjungsan.FItemList(i).FSM_price + ooffjungsan.FItemList(i).FET_price)<>ooffjungsan.FItemList(i).Ftot_jungsanprice then %>
    <br><font color="red"><%= FormatNumber(ooffjungsan.FItemList(i).FTW_price + ooffjungsan.FItemList(i).FUW_price + ooffjungsan.FItemList(i).FCM_price + ooffjungsan.FItemList(i).FOM_price + ooffjungsan.FItemList(i).FSM_price + ooffjungsan.FItemList(i).FET_price,0) %></font>
    <% end if %>
    <% if (ooffjungsan.FItemList(i).Fwaitsum + ooffjungsan.FItemList(i).Ffixedthissum + ooffjungsan.FItemList(i).Ffixednextsum + ooffjungsan.FItemList(i).Fipkumsum)<>ooffjungsan.FItemList(i).Ftot_jungsanprice then %>
    <br><font color="red"><%= FormatNumber(ooffjungsan.FItemList(i).Fwaitsum + ooffjungsan.FItemList(i).Ffixedthissum + ooffjungsan.FItemList(i).Ffixednextsum + ooffjungsan.FItemList(i).Fipkumsum,0) %></font>
    <% end if %>
    
    
    </td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Fwaitsum,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ffixedthissum ,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Ffixednextsum,0) %></td>
    <td align="right"><%= FormatNumber(ooffjungsan.FItemList(i).Fipkumsum,0) %></td>
</tr>
<% if (subtotalFlag) then %>
<tr bgcolor="#DDDDDD" align="right">
    <td align="center"><%= ooffjungsan.FItemList(i).FYYYYMM %></td>
    <td align="center">합계</td>
    <td><%= FormatNumber(sub_TW_price,0) %></td>      
    <td><%= FormatNumber(sub_UW_price,0) %></td>      
    <td><%= FormatNumber(sub_OM_price,0) %></td>      
    <td><%= FormatNumber(sub_SM_price,0) %></td>      
    <td><%= FormatNumber(sub_CM_price,0) %></td>      
    <td><%= FormatNumber(sub_ET_price,0) %></td>      
    <td><%= FormatNumber(sub_jungsanprice,0) %></td>  
    <td><%= FormatNumber(sub_waitsum,0) %></td>       
    <td><%= FormatNumber(sub_fixedthissum,0) %></td>  
    <td><%= FormatNumber(sub_fixednextsum,0) %></td>  
    <td><%= FormatNumber(sub_ipkumsum,0) %></td>      
</tr>
<%
ttl_TW_price        = ttl_TW_price         + sub_TW_price
ttl_UW_price        = ttl_UW_price         + sub_UW_price
ttl_CM_price        = ttl_CM_price         + sub_CM_price
ttl_OM_price        = ttl_OM_price         + sub_OM_price
ttl_SM_price        = ttl_SM_price         + sub_SM_price
ttl_ET_price        = ttl_ET_price         + sub_ET_price
ttl_jungsanprice    = ttl_jungsanprice     + sub_jungsanprice
ttl_waitsum         = ttl_waitsum          + sub_waitsum 
ttl_fixedthissum    = ttl_fixedthissum     + sub_fixedthissum
ttl_fixednextsum    = ttl_fixednextsum     + sub_fixednextsum
ttl_ipkumsum        = ttl_ipkumsum         + sub_ipkumsum

sub_TW_price    = 0
sub_UW_price    = 0
sub_CM_price    = 0
sub_OM_price    = 0
sub_SM_price    = 0
sub_ET_price    = 0
sub_jungsanprice= 0
sub_waitsum     = 0
sub_fixedthissum= 0
sub_fixednextsum= 0
sub_ipkumsum    = 0

end if 

next
%>
<tr bgcolor="#FFFFFF" height=24>
    <td align="center">Total</td>
    <td align="center"></td>
    <td><%= FormatNumber(ttl_TW_price,0) %></td>
    <td><%= FormatNumber(ttl_UW_price,0) %></td>
    <td><%= FormatNumber(ttl_CM_price,0) %></td>
    <td><%= FormatNumber(ttl_OM_price,0) %></td>
    <td><%= FormatNumber(ttl_SM_price,0) %></td>
    <td><%= FormatNumber(ttl_ET_price,0) %></td>
    <td><%= FormatNumber(ttl_jungsanprice,0) %></td>    
    <td><%= FormatNumber(ttl_waitsum,0) %></td>         
    <td><%= FormatNumber(ttl_fixedthissum,0) %></td>    
    <td><%= FormatNumber(ttl_fixednextsum,0) %></td>    
    <td><%= FormatNumber(ttl_ipkumsum,0) %></td> 
</tr>
<% end if %>
</table>

<%
set ooffjungsan = Nothing
%>	

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->