<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls_logics.asp" -->
<%
dim idx
idx = request("idx")

dim oblaju
set oblaju = new CBalju
oblaju.FPageSize=50
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="새로고침" onClick="javascript:document.location.reload();">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

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
	SubTenBaljucount  = SubTenBaljucount +  oblaju.FBaljumasterList(i).FTenBaljucount
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
		<% if CStr(oblaju.FBaljumasterList(i).FBaljuID)=CStr(idx) then %>
		<td><b><font color="#3333AA"><%= oblaju.FBaljumasterList(i).FBaljuID %></font></b></td>
		<% else %>
		<td><%= oblaju.FBaljumasterList(i).FBaljuID %></td>
		<% end if %>
		<td align="left"><a href="?idx=<%= oblaju.FBaljumasterList(i).FBaljuID %>"><%= oblaju.FBaljumasterList(i).FBaljudate %></a></td>
		<td><%= oblaju.FBaljumasterList(i).FTotalBaljucount %></td>
		<td><%= oblaju.FBaljumasterList(i).FUpchecount %></td>
		<td><%= oblaju.FBaljumasterList(i).FTenBaljucount %></td>

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
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
