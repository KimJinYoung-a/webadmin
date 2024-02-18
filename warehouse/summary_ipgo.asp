<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dblogicsopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->

<%

dim startdt, enddt

startdt = Left(dateadd("m",-2,now),10)
enddt = Left(now,10)


dim rackipgo
set rackipgo = new CRackIpgo
rackipgo.FCurrPage = 1
rackipgo.Fpagesize= 200
'rackipgo.FRectExecuteDtStart = "2006-03-01"
'rackipgo.FRectExecuteDtEnd   = toDate
rackipgo.GetRackIpgoList

dim rackjumun
set rackjumun = new CRackIpgo
rackjumun.FCurrPage = 1
rackjumun.Fpagesize= 200
'rackjumun.FRectExecuteDtStart = "2006-03-01"
'rackipgo.FRectExecuteDtEnd   = toDate
rackjumun.GetRackJumunList

dim i, j, dt
dim sumstate0, sumstate1, sumstate5, sumstate7, sumstate8
dim sumrackipgo_y, sumrackipgo_n

sumstate0 = 0
sumstate1 = 0
sumstate5 = 0
sumstate7 = 0
sumstate8 = 0
sumrackipgo_y = 0
sumrackipgo_n = 0

%>
<script language='javascript'>
function GotoIpgoList(ipgodate,rackipgoyn){
        var yyyy, mm, dd
        yyyy = ipgodate.substring(0, 4);
        mm = ipgodate.substring(5, 7);
        dd = ipgodate.substring(8, 10);
	window.open("/admin/newstorage/ipgolist.asp?research=on&ipgocheck=on&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=" + dd + "&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd + "&rackipgoyn=" + rackipgoyn + "&returnyn=N","GotoIpgoList","width=1000,scrollbars=yes");
}

function GotoOrderList(orderdate,statecd){
        var yyyy, mm, dd
        yyyy = orderdate.substring(0, 4);
        mm = orderdate.substring(5, 7);
        dd = orderdate.substring(8, 10);
	window.open("/admin/newstorage/orderlist.asp?research=on&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=" + dd + "&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd + "&statecd=" + statecd,"GotoOrderList","width=1000,scrollbars=yes");
}


function PopUpcheInfo(v){
	window.open("/admin/lib/popupchebrandinfo.asp?designer=" + v,"popupcheinfo","width=640,height=580,scrollbars=yes,resizabled=yes");
}
</script>
<style>
.listSep {
	border-top:0px #CCCCCC solid; height:1px; margin:0; padding:0;
}
</style>
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
	<form name="frm" method="get" action="">
    <input type="hidden" name="research" value="on">
    <input type="hidden" name="page" value="">
    </form>
	<tr>
		<td class="listSep" colspan="15" bgcolor="#AAAAAA" style="border-top:1px"></td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width=80>기준일시</td>
    	<td width=60>주문접수</td>
    	<td width=60>주문확인</td>
    	<td width=60>배송준비</td>
    	<td width=60>출고완료</td>
    	<td width=60>도착완료</td>
        <td width=30></td>
    	<td width=60>입고완료</td>
    	<td width=30></td>
    	<td align=right></td>
    </tr>
<%
dt = enddt
i = 0
j = 0
%>
<% do until (dt = startdt) %>
    <%
    do until (dt >= rackipgo.FItemList(i).Fexecutedt)
            i = i + 1

            if (rackipgo.FResultcount <= i) then
                    i = rackipgo.FResultcount - 1
                    exit do
            end if
    loop

    do until (dt >= rackjumun.FItemList(j).Fexecutedt)
            j = j + 1

            if (rackjumun.FResultcount <= j) then
                    j = rackjumun.FResultcount - 1
                    exit do
            end if
    loop
    %>
<% if (((dt = rackjumun.FItemList(j).Fexecutedt) and (rackjumun.FItemList(j).Ftotalnotfinishcount > 0)) or ((dt = rackipgo.FItemList(i).Fexecutedt) and (rackipgo.FItemList(i).Frackipgo_n > 0))) then %>
    <tr>
		<td class="listSep" colspan="15" bgcolor="#AAAAAA" style="border-top:1px"></td>
	</tr>
	<tr align="center" bgcolor="FFFFFF">
		<td><%= dt %></td>
	<% if (dt = rackjumun.FItemList(j).Fexecutedt) then %>
        <%
        sumstate0 = sumstate0 + rackjumun.FItemList(j).Fstatecd0
        sumstate1 = sumstate1 + rackjumun.FItemList(j).Fstatecd1
        sumstate5 = sumstate5 + rackjumun.FItemList(j).Fstatecd5
        sumstate7 = sumstate7 + rackjumun.FItemList(j).Fstatecd7
        sumstate8 = sumstate8 + rackjumun.FItemList(j).Fstatecd8
        %>
    	<td><a href="javascript:GotoOrderList('<%= rackjumun.FItemList(j).Fexecutedt %>','0');"><%= rackjumun.FItemList(j).Fstatecd0 %></td>
    	<td><a href="javascript:GotoOrderList('<%= rackjumun.FItemList(j).Fexecutedt %>','1');"><%= rackjumun.FItemList(j).Fstatecd1 %></td>
    	<td><a href="javascript:GotoOrderList('<%= rackjumun.FItemList(j).Fexecutedt %>','5');"><%= rackjumun.FItemList(j).Fstatecd5 %></td>
    	<td><a href="javascript:GotoOrderList('<%= rackjumun.FItemList(j).Fexecutedt %>','7');"><%= rackjumun.FItemList(j).Fstatecd7 %></td>
    	<td><a href="javascript:GotoOrderList('<%= rackjumun.FItemList(j).Fexecutedt %>','8');"><%= rackjumun.FItemList(j).Fstatecd8 %></td>
    	<td></td>
	<% else %>
    	<td>0</td>
    	<td>0</td>
    	<td>0</td>
    	<td>0</td>
		<td>0</td>
    	<td></td>
	<% end if %>
	<% if (dt = rackipgo.FItemList(i).Fexecutedt) then %>
        <%
        sumrackipgo_y = sumrackipgo_y + rackipgo.FItemList(i).Frackipgo_y
        sumrackipgo_n = sumrackipgo_n + rackipgo.FItemList(i).Frackipgo_n
        %>
    	<td><b><a href="javascript:GotoIpgoList('<%= rackipgo.FItemList(i).Fexecutedt %>','');"><%= rackipgo.FItemList(i).Frackipgo_n + rackipgo.FItemList(i).Frackipgo_y %></b></a></td>
    	<td></td>
    	<td></td>
	<% else %>
    	<td>0</td>
    	<td>0</td>
    	<td></td>
    	<td align=right>-</td>
    	<td></td>
	<% end if %>
    </tr>
<% end if %>
<% dt = Left((dateadd("d",-1,dt)), 10) %>
<% loop %>
    <tr>
    	<td class="listSep" colspan="15" bgcolor="#AAAAAA" style="border-top:1px"></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>합계</td>
    	<td><%= sumstate0 %></td>
    	<td><%= sumstate1 %></td>
    	<td><%= sumstate5 %></td>
    	<td><%= sumstate7 %></td>
    	<td><%= sumstate8 %></td>
        <td></td>
    	<td><%= sumrackipgo_n + sumrackipgo_y %></td>
    	<td></td>
    	<td></td>
    </tr>
    <tr>
    	<td class="listSep" colspan="15" bgcolor="#AAAAAA" style="border-top:1px"></td>
    </tr>
</table>



<%

set rackipgo = Nothing

%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dblogicsclose.asp" -->
