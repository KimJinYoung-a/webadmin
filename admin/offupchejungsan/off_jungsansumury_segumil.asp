<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
dim yyyy1,mm1, yyyy2, mm2, chkdate
dim research, page
dim taxtype

yyyy1       = request("yyyy1")
mm1         = request("mm1")
yyyy2       = request("yyyy2")
mm2         = request("mm2")
chkdate     = request("chkdate")
research    = request("research")
page        = request("page")
taxtype     = request("taxtype")

if (research="") and (chkdate="") then chkdate="on"
if (page="") then page=1

dim stdt, eddt, StartYYYYMMDD, EndYYYYMMDD
if (yyyy1="") then
	stdt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(stdt),4)
	mm1 = Mid(CStr(stdt),6,2)
	
	eddt = dateadd("d",dateserial(year(Now),month(now)+1,1),-1)
	yyyy2 = Left(CStr(eddt),4)
	mm2 = Mid(CStr(eddt),6,2)
end if


StartYYYYMMDD = yyyy1 + "-" + mm1 + "-" +"01"
EndYYYYMMDD   = Left(CStr(dateserial(yyyy2,mm2+1,1)),10)


dim ooffjungsan
set ooffjungsan = new COffJungsan
ooffjungsan.FPageSize   = 100
ooffjungsan.FCurrPage = page
if (chkdate="on") then
    ooffjungsan.FRectStartYYYYMMDD = StartYYYYMMDD
    ooffjungsan.FRectEndYYYYMMDD   = EndYYYYMMDD
end if

ooffjungsan.FRectTaxType = taxtype
ooffjungsan.GetOffJungsanSummaryBySegumDate


dim i
dim soge
dim tot_jungsansum_susi, tot_jungsansum_31date, tot_jungsansum_15date, tot_jungsansum_etcdate, tot_ewol_jungsansum, tot_tot_jungsanprice, tot_soge, tot_fixedsum, tot_ipkumsum            

soge = 0

tot_jungsansum_susi     = 0
tot_jungsansum_31date   = 0
tot_jungsansum_15date   = 0
tot_jungsansum_etcdate  = 0
tot_ewol_jungsansum     = 0
tot_tot_jungsanprice    = 0
tot_soge                = 0
tot_fixedsum            = 0
tot_ipkumsum            = 0

%>
<script language='javascript'>
function popOffJungsanList(taxregdate,isusual,jungsandate,isipkumfinish){
    var param = 'pop_off_jungsanlist.asp?dategubun=Tax&taxregdate=' + taxregdate + '&isusual=' + isusual + '&jungsandate=' + jungsandate + '&isipkumfinish=' + isipkumfinish;
    var popwin = window.open(param,'pop_off_jungsanlist','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
	        <input type="checkbox" name="chkdate" <% if chkdate="on" then response.write "checked" %> >
        	&nbsp;기간검색 : <% DrawYMYMBox yyyy1,mm1, yyyy2,mm2 %> (계산서 발행월)
        	&nbsp;&nbsp;
        	
        	과세구분 : 
        	<select name="taxtype" >
        	<option value="">전체
        	<option value="01" <%= chkIIF(taxtype="01","selected","") %> >과세
        	<option value="02" <%= chkIIF(taxtype="02","selected","") %> >면세
        	<option value="03" <%= chkIIF(taxtype="03","selected","") %> >원천
        	</select>
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form> 
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	       정상발행 : 세금계산서 발행월이 정산월과 같거나 없을때(원천징수)<br>
	       이월발행 : 세금계산서 발행월이 정산월 이후 일때
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2" width="100">발행월일</td>
		<td colspan="5">정상발행금액</td>
		<td rowspan="2" width="100">이월발행금액</td>
		<td rowspan="2">합계</td>
		<td colspan="2">입금진행여부</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">수시</td>
		<td width="100">15일</td>
		<td width="100">말일</td>
		<td width="100">미지정</td>
		<td width="100">소계</td>
		<td width="100">입금이전</td>
		<td width="100">입금완료</td>
	</tr>
	<% if ooffjungsan.FresultCount<1 then %>
	<tr align="center" bgcolor="#FFFFFF">
	    <td colspan="10" align="center">[검색 결과가 없습니다.]</td>
	</tr>
	<% else %>
	<% for i=0 to ooffjungsan.FresultCount-1 %>
	<%
	    soge = ooffjungsan.FItemList(i).Fjungsansum_susi + ooffjungsan.FItemList(i).Fjungsansum_31date + ooffjungsan.FItemList(i).Fjungsansum_15date + ooffjungsan.FItemList(i).Fjungsansum_etcdate
        
        tot_jungsansum_susi     = tot_jungsansum_susi + ooffjungsan.FItemList(i).Fjungsansum_susi
        tot_jungsansum_31date   = tot_jungsansum_31date + ooffjungsan.FItemList(i).Fjungsansum_31date
        tot_jungsansum_15date   = tot_jungsansum_15date + ooffjungsan.FItemList(i).Fjungsansum_15date
        tot_jungsansum_etcdate  = tot_jungsansum_etcdate + ooffjungsan.FItemList(i).Fjungsansum_etcdate
        tot_ewol_jungsansum     = tot_ewol_jungsansum + ooffjungsan.FItemList(i).Fewol_jungsansum
        tot_tot_jungsanprice    = tot_tot_jungsanprice + ooffjungsan.FItemList(i).Ftot_jungsanprice
        tot_soge                = tot_soge + soge
        tot_fixedsum            = tot_fixedsum + ooffjungsan.FItemList(i).Ffixedsum
        tot_ipkumsum            = tot_ipkumsum + ooffjungsan.FItemList(i).Fipkumsum
	%>
	<tr align="right" bgcolor="#FFFFFF">
		<td align="center"><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','','','A')"><%= ooffjungsan.FItemList(i).Ftaxregdate %></a></td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','Y','수시','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Fjungsansum_susi,0) %></a></td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','Y','15일','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Fjungsansum_15date,0) %></a></td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','Y','말일','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Fjungsansum_31date,0) %></a></td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','Y','NULL','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Fjungsansum_etcdate,0) %></a></td>
		<td>
		    <a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','Y','','A')"><%= FormatNumber(soge,0) %></a>
		</td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','N','','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Fewol_jungsansum,0) %></a></td>
		
		<td>
		    <a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','','','A')"><%= FormatNumber(ooffjungsan.FItemList(i).Ftot_jungsanprice,0) %></a>
		    <% if ooffjungsan.FItemList(i).Ftot_jungsanprice<>(soge + ooffjungsan.FItemList(i).Fewol_jungsansum) then %>
		    <br><font color="red"><%= FormatNumber(soge + ooffjungsan.FItemList(i).Fewol_jungsansum,0) %></font>
		    <% end if %>
		    
		    <% if ooffjungsan.FItemList(i).Ftot_jungsanprice<>(ooffjungsan.FItemList(i).Fipkumsum + ooffjungsan.FItemList(i).Ffixedsum) then %>
		    <br><font color="blue"><%= FormatNumber(ooffjungsan.FItemList(i).Fipkumsum + ooffjungsan.FItemList(i).Ffixedsum,0) %></font>
		    <% end if %>
		</td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','','','N')"><%= FormatNumber(ooffjungsan.FItemList(i).Ffixedsum,0) %></a></td>
		<td><a href="javascript:popOffJungsanList('<%= ooffjungsan.FItemList(i).Ftaxregdate %>','','','Y')"><%= FormatNumber(ooffjungsan.FItemList(i).Fipkumsum,0) %></a></td>
	</tr>
	<% next %>
	<% end if %>
	
	<tr align="right" bgcolor="#DDDDDD">
		<td align="center" >Total</td>
		<td><%= FormatNumber(tot_jungsansum_susi,0) %></td>   
		<td><%= FormatNumber(tot_jungsansum_15date,0) %></td> 
		<td><%= FormatNumber(tot_jungsansum_31date,0) %></td> 
		<td><%= FormatNumber(tot_jungsansum_etcdate,0) %></td>
		<td><%= FormatNumber(tot_soge,0) %></td>   
		<td><%= FormatNumber(tot_ewol_jungsansum,0) %></td>  
		<td><%= FormatNumber(tot_tot_jungsanprice,0) %></td>              
		<td><%= FormatNumber(tot_fixedsum,0) %></td>          
		<td><%= FormatNumber(tot_ipkumsum,0) %></td>          
	</tr>
</table>
<%
set ooffjungsan = Nothing
%>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->