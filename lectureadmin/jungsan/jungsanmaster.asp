<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim id
id = RequestCheckvar(request("id"),10)
dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FRectId = id
ojungsan.FRectdesigner = session("ssBctID")
ojungsan.JungsanMasterList

if ojungsan.FresultCount <1 then
	dbget.close()	:	response.End
end if

dim rd_state
rd_state = ojungsan.FItemList(0).Ffinishflag
%>
<script language='javascript'>
function savestate(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret){
		frm.submit();
	}
}
</script>
<br>
<!--
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="popshowdetail.asp?menupos=<%= menupos %>&id=<%= id %>">상세내역&gt;&gt;</a></td>
</tr>
</table>
-->
<!--
<div class="a">[정산 내역이 맞으면 <b>업체확인완료</b>를 누르신후 저장하시기 바랍니다.]</div>
-->
<br>
<div class="a">1.기준정보</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form name="statefrm" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="statechange">
<input type="hidden" name="idx" value="<%= ojungsan.FItemList(0).FId %>">
    <tr >
      <td width="100" bgcolor="#DDDDFF">브랜드ID</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).Fdesignerid %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">정산대상년월</td>
      <td bgcolor="#FFFFFF"><%= ojungsan.FItemList(0).FYYYYMM %></td>
    </tr>
    <tr >
      <td width="100" bgcolor="#DDDDFF">현재상태</td>
      <td bgcolor="#FFFFFF">
		<input type="radio" name="rd_state" value="1" <% if rd_state="1" then response.write "checked" %> >업체확인대기
		<input type="radio" name="rd_state" value="2" <% if rd_state<>"1" then response.write "checked" %> >업체확인완료
		<input type="button" value="저장" onclick="savestate(statefrm);" <% if rd_state<>"1" then response.write "disabled" %> >
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">세금계산서발행일</td>
      <td bgcolor="#FFFFFF">
      	<%= ojungsan.FItemList(0).Ftaxregdate %>
      </td>
    </tr>
    <tr>
      <td width="100" bgcolor="#DDDDFF">입금일</td>
      <td bgcolor="#FFFFFF">
      	<%= ojungsan.FItemList(0).Fipkumdate %>
      </td>
    </tr>
</form>
</table>

<br>
<div class="a"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">2.정산내역</a></div>
<table width="760" cellspacing="1" cellpadding=2 class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
	<td width=100 align=left>구분</td>
	<td width=100>총주문건수</td>
	<td width=100>소비자가총액</td>
	<td width=100>공급가총액</td>
	<td width=70>마진</td>
	<td>기타</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=upche">업체배송</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fub_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fub_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fub_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fub_totalsuplycash/ojungsan.FItemList(0).Fub_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fub_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=maeip">매입내역</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fme_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fme_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fme_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fme_totalsuplycash/ojungsan.FItemList(0).Fme_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fme_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF"><a href="nowjungsandetail.asp?id=<%= ojungsan.FItemList(0).Fid %>&gubun=witaksell">특정온라인내역</a></td>
	<td align=right><%= ojungsan.FItemList(0).Fwi_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fwi_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fwi_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fwi_totalsuplycash/ojungsan.FItemList(0).Fwi_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=center></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fwi_comment) %></td>
</tr>
<!--
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">특정 오프라인</td>
	<td><%= ojungsan.FItemList(0).Fsh_cnt %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsellcash,0) %></td>
	<td><%= FormatNumber(ojungsan.FItemList(0).Fsh_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fsh_totalsellcash<>0 then %>
	<td><%= CLng((1-ojungsan.FItemList(0).Fsh_totalsuplycash/ojungsan.FItemList(0).Fsh_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td>?</td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fsh_comment) %></td>
</tr>
-->
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">특정 기타</td>
	<td align=right><%= ojungsan.FItemList(0).Fet_cnt %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).Fet_totalsuplycash,0) %></td>
	<% if ojungsan.FItemList(0).Fet_totalsellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).Fet_totalsuplycash/ojungsan.FItemList(0).Fet_totalsellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td><%= nl2br(ojungsan.FItemList(0).Fet_comment) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF">총계</td>
	<td></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSellcash,0) %></td>
	<td align=right><%= FormatNumber(ojungsan.FItemList(0).GetTotalSuplycash,0) %></td>
	<% if ojungsan.FItemList(0).GetTotalSellcash<>0 then %>
	<td align=center><%= CLng((1-ojungsan.FItemList(0).GetTotalSuplycash/ojungsan.FItemList(0).GetTotalSellcash)*10000)/100 %> %</td>
	<% else %>
	<td align=right></td>
	<% end if %>
	<td></td>
</tr>
</table>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->