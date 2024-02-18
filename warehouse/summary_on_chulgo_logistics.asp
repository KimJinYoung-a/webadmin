<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_baljucls.asp" -->
<%
dim idx, siteSeq
idx = RequestCheckVar(request("idx"),10)
siteSeq = RequestCheckVar(request("siteSeq"),10)

if (siteSeq="") then siteSeq="10"

dim oblaju
set oblaju = new CBalju
oblaju.FPageSize=50
oblaju.FRectSiteSeq = siteSeq
oblaju.getBaljumasterInfoList


dim i, isdayfinal,predate


dim SubTotalBaljucount, SubUpchecount
dim SubTenBaljucount, SubMibeacount
dim SubWaitcount, SubIpgoCount
dim SubPrintCount, SubPackingCount
dim SubuploadCount, SubCancelCount
dim SubEtcCount

dim TotalBaljucount, Upchecount
dim TenBaljucount, Mibeacount
dim Waitcount, IpgoCount
dim PrintCount, PackingCount
dim uploadCount, CancelCount
dim TotalEtcCount
%>
<script language='javascript'>
function RefreshBaljuMaster(preidx){
	if (confirm('출고지시 목록을 업데이트 하시겠습니까?')){
		document.refFrm.preidx.value=preidx;
		document.refFrm.mode.value="refbaljumasterjustone";
		document.refFrm.submit();
	}
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
		<td align="left">
    		사이트 : 
            <select class="select" name="siteseq">
           <!-- <option value="">SITE 선택 -->
           <option value="10" <%= CHKIIF(siteSeq="10","selected","") %> >텐바이텐</option>
           <option value="30" <%= CHKIIF(siteSeq="30","selected","") %> >유아러걸</option>
           <option value="50" <%= CHKIIF(siteSeq="50","selected","") %> >탐스슈즈</option>
           <option value="99" <%= CHKIIF(siteSeq="99","selected","") %> >아이띵소</option>
           </select>
		</td>
		
		<td rowspan="2" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>
 
				
						    
<table width="100%" align="center" cellpadding="3" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=40>출고지시<br>번호</td>
		<td width=140>출고지시일시</td>
		<td width=50>총출고지시</td>
		<td width=50>업배</td>
		<td width=50>텐배</td>
		<td width=50>상품준비<br>(미입고)</td>
		<td width=50>출고준비<br>(미출력)</td>
		<td width=50>출고준비<br>(출력완)</td>
		<td width=50>취소</td>
		<td width=50>기타</td>
		
		<td width=50>총<br>출고완료</td>
		<td width=50>당일<br>출고완료</td>
		<td width=50>1일<br>지연출고</td>
		<td width=50>2일<br>지연출고</td>
		<td width=50>3일이상<br>지연출고</td>
		<td width=50>완료율</td>
		<td width=50>상품준비<br>(미배)</td>
		<td></td>
	</tr>
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<% for i=0 to oblaju.FResultCount - 1 %>
	<%
	if (predate<>"") and (predate<>Left(oblaju.FBaljumasterList(i).FBaljudate,10)) then
		TotalBaljucount = TotalBaljucount + SubTotalBaljucount
		Upchecount 		= Upchecount + SubUpchecount
		TenBaljucount	= TenBaljucount + SubTenBaljucount
		Mibeacount   	= Mibeacount + SubMibeacount
		Waitcount		= Waitcount + SubWaitcount
		IpgoCount      	= IpgoCount + SubIpgoCount
		PrintCount		= PrintCount + SubPrintCount
		PackingCount   	= PackingCount + SubPackingCount
		uploadCount		= uploadCount + SubuploadCount
		CancelCount    	= CancelCount + SubCancelCount
		TotalEtcCount 	= TotalEtcCount + SubEtcCount
	%>
	<tr align=center bgcolor="#DDDDDD" >
		<td ></td>
		<td ></td>
		<td ><%= SubTotalBaljucount %></td>
		<td ><%= SubUpchecount %></td>
		<td ><%= SubTenBaljucount %></td>

		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=0" target="_blank"><%= SubWaitcount %></a></td>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=3" target="_blank"><%= SubIpgoCount %></a></td>
		<% if SubPrintCount>0 then %>
		<td ><b><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=5" target="_blank"><%= SubPrintCount %></a></b></td>
		<% else %>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=5" target="_blank"><%= SubPrintCount %></a></td>
		<% end if %>
		<td><%= SubCancelCount %></td>
		<td><%= SubEtcCount %></td>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_list.asp?baljudate=<%= predate %>&baljuflag=7" target="_blank"><%= SubPackingCount %></a></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td >
		<% if SubTenBaljucount-SubCancelCount<>0 then %>
			<% if SubPackingCount =SubTenBaljucount-SubCancelCount-SubEtcCount then %>
			<b><font color=red><%= CLng((SubPackingCount )/(SubTenBaljucount-SubCancelCount-SubEtcCount)*100*100)/100 %>%</font></b>
			<% else %>
			<%= CLng((SubPackingCount )/(SubTenBaljucount-SubCancelCount-SubEtcCount)*100*100)/100 %>%
			<% end if %>
		<% end if %>
		</td>
        <% if SubMibeacount>0 then %>
		<td ><b><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=2" target="_blank"><%= SubMibeacount %></a></b></td>
		<% else %>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=2" target="_blank"><%= SubMibeacount %></a></td>
		<% end if %>
		
		<td></td>
		
	</tr>
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<%
		SubTotalBaljucount =0
		SubUpchecount  =0
		SubTenBaljucount  =0
		SubMibeacount  =0
		SubWaitcount  =0
		SubIpgoCount  =0
		SubPrintCount  =0
		SubPackingCount  =0
		SubuploadCount  =0
		SubCancelCount =0
		SubEtcCount = 0
	end if
	%>
	<%
	predate = Left(oblaju.FBaljumasterList(i).FBaljudate,10)
	%>
	<%
	SubTotalBaljucount = SubTotalBaljucount + oblaju.FBaljumasterList(i).FTotalBaljucount
	SubUpchecount  = SubUpchecount +  oblaju.FBaljumasterList(i).FUpchecount
	SubTenBaljucount  = SubTenBaljucount +  oblaju.FBaljumasterList(i).FLocalBaljucount
	SubMibeacount  = SubMibeacount +  oblaju.FBaljumasterList(i).FMibeacount
	SubWaitcount  = SubWaitcount +  oblaju.FBaljumasterList(i).FWaitcount
	SubIpgoCount  = SubIpgoCount  +  oblaju.FBaljumasterList(i).FIpgoCount
	SubPrintCount  = SubPrintCount +  oblaju.FBaljumasterList(i).FPrintCount
	SubPackingCount  = SubPackingCount  +  oblaju.FBaljumasterList(i).FPackingCount
	SubuploadCount  = SubuploadCount  +  oblaju.FBaljumasterList(i).FuploadCount
	SubCancelCount = SubCancelCount  +  oblaju.FBaljumasterList(i).FCancelCount
	SubEtcCount = SubEtcCount  +  oblaju.FBaljumasterList(i).FEtcCount
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<% if CStr(oblaju.FBaljumasterList(i).FSiteBaljuID)=CStr(idx) then %>
		<td><b><font color="#3333AA"><%= oblaju.FBaljumasterList(i).FSiteBaljuID %></font></b></td>
		<% else %>
		<td><%= oblaju.FBaljumasterList(i).FSiteBaljuID %></td>
		<% end if %>
		<td align="left"><a href="?idx=<%= oblaju.FBaljumasterList(i).FSiteBaljuID %>"><%= oblaju.FBaljumasterList(i).FBaljudate %></a></td>
		<td><%= oblaju.FBaljumasterList(i).FTotalBaljucount %></td>
		<td><%= oblaju.FBaljumasterList(i).FUpchecount %></td>
		<td><%= oblaju.FBaljumasterList(i).FLocalBaljucount %></td>

		<td><%= oblaju.FBaljumasterList(i).FWaitcount %></td>
		<td><%= oblaju.FBaljumasterList(i).FIpgoCount %></td>
		<td><%= oblaju.FBaljumasterList(i).FPrintCount %></td>
		<td><%= oblaju.FBaljumasterList(i).FCancelCount  %></td>
		<td><%= oblaju.FBaljumasterList(i).FEtcCount  %></td>
		<td><%= oblaju.FBaljumasterList(i).FPackingCount %></td>
		<td><%= oblaju.FBaljumasterList(i).Fdelay0chulgocnt %></td>
		<td><%= oblaju.FBaljumasterList(i).Fdelay1chulgocnt %></td>
		<td><%= oblaju.FBaljumasterList(i).Fdelay2chulgocnt %></td>
		<td><%= oblaju.FBaljumasterList(i).Fdelay3chulgocnt %></td>
		<td ></td>
        <td><%= oblaju.FBaljumasterList(i).FMibeacount %></td>
        <td></td>
        
	</tr>
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<% next %>
	<tr align=center  bgcolor="#DDDDDD" >
		<td ></td>
		<td ></td>
		<td ><%= SubTotalBaljucount %></td>
		<td ><%= SubUpchecount %></td>
		<td ><%= SubTenBaljucount %></td>
		
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=0" target="_blank"><%= SubWaitcount %></a></td>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=3" target="_blank"><%= SubIpgoCount %></a></td>
		<% if SubPrintCount>0 then %>
		<td ><b><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=5" target="_blank"><%= SubPrintCount %></a></b></td>
		<% else %>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=5" target="_blank"><%= SubPrintCount %></a></td>
		<% end if %>
		<td><%= SubCancelCount %></td>
		<td><%= SubEtcCount %></td>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_list.asp?baljudate=<%= predate %>&baljuflag=7" target="_blank"><%= SubPackingCount %></a></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td >
		<% if SubTenBaljucount-SubCancelCount<>0 then %>
			<% if SubPackingCount =SubTenBaljucount-SubCancelCount-SubEtcCount then %>
			<b><font color=red><%= CLng((SubPackingCount )/(SubTenBaljucount-SubCancelCount-SubEtcCount)*100*100)/100 %>%</font></b>
			<% else %>
			<%= CLng((SubPackingCount )/(SubTenBaljucount-SubCancelCount-SubEtcCount)*100*100)/100 %>%
			<% end if %>
		<% end if %>
		</td>


        <% if SubMibeacount>0 then %>
		<td ><b><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=2" target="_blank"><%= SubMibeacount %></a></b></td>
		<% else %>
		<td ><a href="http://logics.10x10.co.kr/m_chulgo_packinglist.asp?baljudate=<%= predate %>&baljuflag=2" target="_blank"><%= SubMibeacount %></a></td>
		<% end if %>
		
		<td></td>
	</tr>
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
	<%
	TotalBaljucount = TotalBaljucount + SubTotalBaljucount
	Upchecount 		= Upchecount + SubUpchecount
	TenBaljucount	= TenBaljucount + SubTenBaljucount
	Mibeacount   	= Mibeacount + SubMibeacount
	Waitcount		= Waitcount + SubWaitcount
	IpgoCount      	= IpgoCount + SubIpgoCount
	PrintCount		= PrintCount + SubPrintCount
	PackingCount   	= PackingCount + SubPackingCount
	uploadCount		= uploadCount + SubuploadCount
	CancelCount    	= CancelCount + SubCancelCount
	TotalEtcCount 	= TotalEtcCount + SubEtcCount
	%>
	<tr align=center  bgcolor="#EEEE22" >
		<td >Total</td>
		<td ></td>
		<td ><%= TotalBaljucount %></td>
		<td ><%= Upchecount %></td>
		<td ><%= TenBaljucount %></td>
		<td ><%= Waitcount %></td>
		<td ><%= IpgoCount %></td>
		<td ><%= PrintCount %></td>
		<td><%= CancelCount %></td>
		<td><%= TotalEtcCount %></td>
		<td ><%= PackingCount %></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td >
		<% if TenBaljucount-CancelCount<>0 then %>
			<%= CLng((PackingCount )/(TenBaljucount-CancelCount-TotalEtcCount)*100*100)/100 %>%
		<% end if %>
		</td>
        <td ><%= Mibeacount %></td>
        
        <td></td>
	</tr>
	<tr>
		<td height="1" colspan="20" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
</table>
									
			

<%
set oblaju = Nothing
%>
<form name=refFrm method=post action="/action/actrefresh.asp">
<input type=hidden name=mode value="refbaljumasterjustone">
<input type=hidden name=idx value="">
<input type=hidden name=preidx value="">
</form>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
