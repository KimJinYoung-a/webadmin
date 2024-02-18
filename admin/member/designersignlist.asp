<%@ language=vbscript %>
<% option explicit %>
<%
response.write "사용중지 메뉴입니다. - 관리자 문의 요망"
dbget.close()	:	response.End
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim rectDesigner
rectDesigner = request("rectDesigner")

dim rectMaeip
rectMaeip = request("rectMaeip")

class CSocMarginSum
	public Fsocid
	public Fmarginsumstr

	Private Sub Class_Initialize()
                '
	End Sub

	Private Sub Class_Terminate()
                '
	End Sub
end class

'==============================================================================
dim sqlStr, i, j, k, tmp

'상품별마진
sqlStr = " select c.userid, IsNull(T.mwdiv, '') as mwdiv, IsNull(T.margine,0) as margine, IsNull(T.cnt, 0) as cnt "
sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_c c "
sqlStr = sqlStr + " left join "
sqlStr = sqlStr + " ( "
sqlStr = sqlStr + " select makerid, "
sqlStr = sqlStr + " ( "
sqlStr = sqlStr + " case deliverytype when '2' then 'U' "
sqlStr = sqlStr + " 	when '5' then 'U' "
sqlStr = sqlStr + " else "
sqlStr = sqlStr + " 	mwdiv "
sqlStr = sqlStr + " end "
sqlStr = sqlStr + " ) as mwdiv "
'sqlStr = sqlStr + " ,(100-buycash/sellcash*100) as margine, count(itemid) as cnt "
sqlStr = sqlStr + " ,(100-orgsuplycash/orgprice*100) as margine, count(itemid) as cnt "
sqlStr = sqlStr + " from [db_item].[dbo].tbl_item "
sqlStr = sqlStr + " where orgprice<>0 "
sqlStr = sqlStr + " and isusing='Y' "
sqlStr = sqlStr + " group by makerid,( "
sqlStr = sqlStr + " case deliverytype when '2' then 'U' "
sqlStr = sqlStr + " 	when '5' then 'U' "
sqlStr = sqlStr + " else "
sqlStr = sqlStr + " 	mwdiv "
sqlStr = sqlStr + " end "
sqlStr = sqlStr + " ),(100-orgsuplycash/orgprice*100) "
sqlStr = sqlStr + " ) as T on c.userid=T.makerid "
sqlStr = sqlStr + " where c.userdiv<21 "
sqlStr = sqlStr + " and c.isusing='Y' "
if (rectDesigner <> "") then
        sqlStr = sqlStr + " and c.userid = '" + rectDesigner + "' "
end if
sqlStr = sqlStr + " order by c.userid, T.cnt desc "
rsget.Open sqlStr,dbget,1

dim socmarginsum()
redim preserve socmarginsum(rsget.RecordCount)
i = 0
do until rsget.Eof
        if (rsget("mwdiv") = "M") then
                tmp = "매입"
        elseif (rsget("mwdiv") = "W") then
                tmp = "위탁"
        elseif (rsget("mwdiv") = "U") then
                tmp = "업체"
        else
                tmp = "????"
        end if

        if (i = 0) then
                j = 0
                set socmarginsum(j) = new CSocMarginSum
                socmarginsum(j).Fsocid = rsget("userid")
                socmarginsum(j).Fmarginsumstr = CStr(rsget("cnt")) + "(" + tmp + "," + CStr(rsget("margine")) + "%)"
        elseif (rsget("userid") <> socmarginsum(j).Fsocid) then
                j = j + 1
                set socmarginsum(j) = new CSocMarginSum
                socmarginsum(j).Fsocid = rsget("userid")
                socmarginsum(j).Fmarginsumstr = CStr(rsget("cnt")) + "(" + tmp + "," + CStr(rsget("margine")) + "%)"
        else
                if (socmarginsum(j).Fmarginsumstr = "") then
                        socmarginsum(j).Fmarginsumstr = CStr(rsget("cnt")) + "(" + tmp + "," + CStr(rsget("margine")) + "%)"
                else
                        socmarginsum(j).Fmarginsumstr = socmarginsum(j).Fmarginsumstr + ", " + CStr(rsget("cnt")) + "(" + tmp + "," + CStr(rsget("margine")) + "%)"
                end if
        end if
	i=i+1
	rsget.MoveNext
loop
rsget.Close


'==============================================================================

'업체별마진
sqlStr = " select a.*, b.*, c.*, d.*, e.*, f.*, g.* "
sqlStr = sqlStr + " from ( "
' '10x10'
sqlStr = sqlStr + "         select c.userid as makerid0,c.maeipdiv as chargediv0,c.defaultmargine as defaultmargine0,c.defaultmargine as defaultsuplymargin0,c.isusing "
sqlStr = sqlStr + "         from [db_user].[dbo].tbl_user_c c "
sqlStr = sqlStr + "         where c.userdiv < 15 "
sqlStr = sqlStr + "          "
if (rectMaeip <> "") then
        sqlStr = sqlStr + " and c.maeipdiv = '" + rectMaeip + "' "
end if
if (rectDesigner <> "") then
        sqlStr = sqlStr + " and c.userid = '" + rectDesigner + "' "
end if

sqlStr = sqlStr + " ) a left join ( "
' 'streetshop000'
sqlStr = sqlStr + "         select s.makerid as makerid1,s.chargediv as chargediv1,s.defaultmargin as defaultmargine1,s.defaultsuplymargin as defaultsuplymargin1 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop000' "
sqlStr = sqlStr + " ) b on makerid0 = b.makerid1 "

sqlStr = sqlStr + " left join ( "
' 'streetshop001'
sqlStr = sqlStr + "         select s.makerid as makerid2,s.chargediv as chargediv2,s.defaultmargin as defaultmargine2,s.defaultsuplymargin as defaultsuplymargin2 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop001' "
sqlStr = sqlStr + " ) c on makerid0 = c.makerid2 "
sqlStr = sqlStr + " left join ( "
' 'streetshop002'
sqlStr = sqlStr + "         select s.makerid as makerid3,s.chargediv as chargediv3,s.defaultmargin as defaultmargine3,s.defaultsuplymargin as defaultsuplymargin3 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop002' "
sqlStr = sqlStr + " ) d on a.makerid0 = d.makerid3 "
sqlStr = sqlStr + " left join ( "
' 'streetshop003'
sqlStr = sqlStr + "         select s.makerid as makerid4,s.chargediv as chargediv4,s.defaultmargin as defaultmargine4,s.defaultsuplymargin as defaultsuplymargin4 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop003' "
sqlStr = sqlStr + " ) e on a.makerid0 = e.makerid4 "
sqlStr = sqlStr + " left join ( "
' 'streetshop004'
sqlStr = sqlStr + "         select s.makerid as makerid5,s.chargediv as chargediv5,s.defaultmargin as defaultmargine5,s.defaultsuplymargin as defaultsuplymargin5 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop004' "
sqlStr = sqlStr + " ) f on a.makerid0 = f.makerid5 "
sqlStr = sqlStr + " left join ( "
' 'streetshop800'
sqlStr = sqlStr + "         select s.makerid as makerid8,s.chargediv as chargediv8,s.defaultmargin as defaultmargine8,s.defaultsuplymargin as defaultsuplymargin8 "
sqlStr = sqlStr + "         from [db_shop].[dbo].tbl_shop_designer s "
sqlStr = sqlStr + "         where s.shopid='streetshop800' "
sqlStr = sqlStr + " ) g on a.makerid0 = g.makerid8 "
sqlStr = sqlStr + " order by chargediv2 desc , chargediv3 desc, chargediv4 desc, chargediv5 desc, chargediv8 desc"

%>

<script language='javascript'>

function popbrandinfoonly(v){
	window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popupcheinfo","width=640 height=640 scrollbars=yes resizable=yes");
}

function editOffDesinger(shopid,designerid){
	var popwin = window.open("/admin/lib/popshopupcheinfo.asp?shopid=" + shopid + "&designer=" + designerid,"popshopupcheinfo","width=640 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>

<table width="95%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a">
		아이디 <input type="text" name="rectDesigner" value="<%= rectDesigner %>" Maxlength="32" size="16">
		<input type=radio name=rectMaeip value="" <% if (rectMaeip = "") then response.write "checked" end if %>> 전체
		<input type=radio name=rectMaeip value="M" <% if (rectMaeip = "M") then response.write "checked" end if %>> 매입
		<input type=radio name=rectMaeip value="W" <% if (rectMaeip = "W") then response.write "checked" end if %>> 위탁
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table><br>

<table width="95%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#000000>
<tr bgcolor="#DDDDFF">
  <td rowspan=2 align=center>브랜드ID</td>
  <td colspan=3 align=center>텐바이텐</td>
  <td colspan=4 align=center>streetshop000</td>
  <td colspan=4 align=center>streetshop001</td>
  <td colspan=4 align=center>streetshop002</td>
  <td colspan=4 align=center>streetshop003</td>
  <td colspan=4 align=center>streetshop004</td>
  <td colspan=4 align=center>streetshop800</td>
</tr>
<tr bgcolor="#DDDDFF">
  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>

  <td align=center>구분</td>
  <td align=center>마진</td>
  <td align=center>제공</td>
  <td align=center>수정</td>
</tr>
<%
rsget.Open sqlStr,dbget,1
do until rsget.Eof
%>

<% if rsget("isusing")="Y" then %>
<tr bgcolor="#FFFFFF">
<% else %>
<tr bgcolor="#CCCCCC">
<% end if %>
<!-- <td rowspan=2> -->
  <td><%= rsget("makerid0") %></td>
  <td align=center>
        <%
        if (rsget("chargediv0") = "M") then
                response.write "매입"
        elseif (rsget("chargediv0") = "W") then
                response.write "위탁"
        elseif (rsget("chargediv0") = "U") then
                response.write "업체"
        else
                response.write rsget("chargediv0")
        end if
        %>
  </td>
  <td align=center><%= rsget("defaultmargine0") %></td>
  <td align=center><a href="javascript:popbrandinfoonly('<%= rsget("makerid0") %>')">>></a></td>

  <td bgcolor="#EEEEEE" align=center>
        <%
        if (rsget("chargediv1") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv1") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv1") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv1") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv1")
        end if
        %>
  </td>
  <td align=center bgcolor="#EEEEEE"><%= rsget("defaultmargine1") %></td>
  <td align=center bgcolor="#EEEEEE"><%= rsget("defaultsuplymargin1") %></td>
  <td align=center bgcolor="#EEEEEE"><a href="javascript:editOffDesinger('streetshop000','<%= rsget("makerid0") %>')">>></a></td>

  <td align=center>
        <%
        if (rsget("chargediv2") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv2") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv2") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv2") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv2")
        end if
        %>
  </td>
  <td align=center><%= rsget("defaultmargine2") %></td>
  <td align=center><%= rsget("defaultsuplymargin2") %></td>
  <td align=center><a href="javascript:editOffDesinger('streetshop001','<%= rsget("makerid0") %>')">>></a></td>

  <td align=center>
        <%
        if (rsget("chargediv3") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv3") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv3") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv3") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv3")
        end if
        %>
  </td>
  <td align=center><%= rsget("defaultmargine3") %></td>
  <td align=center><%= rsget("defaultsuplymargin3") %></td>
  <td align=center><a href="javascript:editOffDesinger('streetshop002','<%= rsget("makerid0") %>')">>></a></td>

  <td align=center>
        <%
        if (rsget("chargediv4") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv4") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv4") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv4") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv4")
        end if
        %>
  </td>
  <td align=center><%= rsget("defaultmargine4") %></td>
  <td align=center><%= rsget("defaultsuplymargin4") %></td>
  <td align=center><a href="javascript:editOffDesinger('streetshop003','<%= rsget("makerid0") %>')">>></a></td>

  <td align=center>
        <%
        if (rsget("chargediv5") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv5") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv5") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv5") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv5")
        end if
        %>
  </td>
  <td align=center><%= rsget("defaultmargine5") %></td>
  <td align=center><%= rsget("defaultsuplymargin5") %></td>
  <td align=center><a href="javascript:editOffDesinger('streetshop004','<%= rsget("makerid0") %>')">>></a></td>

  <td align=center bgcolor="#EEEEEE">
        <%
        if (rsget("chargediv8") = "2") then
                response.write "텐위"
        elseif (rsget("chargediv8") = "4") then
                response.write "텐매"
        elseif (rsget("chargediv8") = "6") then
                response.write "<font color=red><b>업위</b></font>"
        elseif (rsget("chargediv8") = "8") then
                response.write "업매"
        else
                response.write rsget("chargediv8")
        end if
        %>
  </td>
  <td align=center bgcolor="#EEEEEE"><%= rsget("defaultmargine8") %></td>
  <td align=center bgcolor="#EEEEEE"><%= rsget("defaultsuplymargin8") %></td>
  <td align=center bgcolor="#EEEEEE"><a href="javascript:editOffDesinger('streetshop800','<%= rsget("makerid0") %>')">>></a></td>
</tr>


<!--
<tr bgcolor="#FFFFFF">
  <td colspan=27>
        <%
        for k = 0 to j
                if (socmarginsum(k).Fsocid = rsget("makerid0")) then
                        response.write socmarginsum(k).Fmarginsumstr
                end if
        next
        %>
    &nbsp;
  </td>
</tr>
-->

<%
	rsget.MoveNext
loop
rsget.Close
%>

</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->