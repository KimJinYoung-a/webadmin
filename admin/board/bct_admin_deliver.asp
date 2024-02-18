<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<!-- #include virtual="/lib/classes/noreplyboardcls.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "배송유의사항"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, nextdateStr

dim topn,designer,page

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
page = request("page")
if (page="") then page=1

nowdateStr = CStr(now())

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))


dim i
dim oneboard
set oneboard = new CNoReplyBoard
oneboard.Currentpage= page
''oneboard.deleteyn = "N"
''oneboard.checkflag = "N"

oneboard.listBoard
%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->
<!--
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<tr>
  <td valign="top"> 
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
      <tr valign="top"> 
        <td height="5" class="a" width="315"> 
          <div align="left">조회기간 </div>
        </td>
        <td width="85" class="a" height="5">글쓴이<br>
        </td>
        <td width="100" class="a" height="5">처리여부<br>
        </td>
        <td width="100" class="a" height="5">주문번호</td>
        <td width="48" class="a" height="5">고객명</td>
        <td width="49" class="a" height="5">기타검색</td>
        <td class="a" height="5">&nbsp; </td>
      </tr>
      <tr valign="top"> 
        <td width="315" class="a"> 
          <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        </td>
        <td width="85" class="a"> 
          
        </td>
        <td width="100" class="a"> 
          
        </td>
        <td width="100" class="a"> 
          <input type="text" name="textfield222" size="15">
        </td>
        <td width="48" class="a"> 
          <input type="text" name="textfield22" size="15">
        </td>
        <td width="49" class="a"> 
          
        </td>
        <td class="a"> <img src="/images/search2.gif" width="74" height="22"></td>
      </tr>
    </table>
  </td>
</tr>
</table>
-->

<br>
<table width="720" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
	<td><font color="red"><a href="javascript:AnPopDelivery()">print</a></font></td>
	<td class="a" align="right"><%= CStr(oneboard.ResultCount) + "/" + Cstr(oneboard.FTotalCount) %></td>
</tr>
<tr> 
  <td class="a" width="409"><b><img src="/admin/images/mini_icon.gif" width="17" height="17"> 
    배송유의사항 리스트</b></td>
  <td class="a"> 
    <div align="right"><a href="bct_admin_deliver_write.asp"><img src="/admin/images/write_butten.gif" width="55" height="17" border="0"></a> 
    </div>
  </td>
</tr>
</table>
<table width="680" border="0" cellpadding="0" cellspacing="3" align="center">
<tr> 
	<td height="25" valign="middle" class="top_bg"> 
	  <div align="center"> 
	    <table width="720" border="0" cellpadding="0" cellspacing="0" class="a">
	      <tr> 
	        <td width="35"> 
	          <div align="center">번호</div>
	        </td>
	        <td> 
	          <div align="center">제 목</div>
	        </td>
	        <td width="65"> 
	          <div align="center">사이트</div>
	        </td>
	        <td width="90"> 
	          <div align="center">주문번호</div>
	        </td>
	        <td width="60"> 
	          <div align="center">글쓴이 </div>
	        </td>
	        <td width="60"> 
	          <div align="center">고객명</div>
	        </td>
	        <td width="70"> 
	          <div align="center">날 짜</div>
	        </td>
	        <td width="58"> 
	          <div align="center">처리여부</div>
	        </td>
	      </tr>
	    </table>
	  </div>
	</td>
</tr>
<%
for i=0 to oneboard.resultcount-1
%>
<tr> 
	<td valign="top"> 
	  <div align="center"> 
	    <table width="720" border="0" cellpadding="0" cellspacing="0" class="a">
	      <tr onMouseOver="this.style.backgroundColor='#eeeeee'" onMouseOut="this.style.backgroundColor='#ffffff'" bgcolor="white" height="20" > 
	        <td width="35"> 
	          <div align="center"><%= oneboard.FBoardItem(i).FID %></div>
	        </td>
	        <td><a href="#"><a href="/admin/board/bct_admin_deliver_read.asp?id=<%= oneboard.FBoardItem(i).FID %>"><%= oneboard.FBoardItem(i).Ftitle %></a></a></td>
	        <td width="65"> 
	          <div align="center"><%= oneboard.FBoardItem(i).FSitename %></div>
	        </td>
	        <td width="90"> 
	          <div align="center" class="id"><%= oneboard.FBoardItem(i).FOrderSerial %></div>
	        </td>
	        <td width="60"> 
	          <div align="center"><%= oneboard.FBoardItem(i).Fwriter %></div>
	        </td>
	        <td width="60"> 
	          <div align="center" class="a"><%= oneboard.FBoardItem(i).FBuyName %></div>
	        </td>
	        <td width="70"> 
	          <div align="center"><%= oneboard.FBoardItem(i).FMatchDate %></div>
	        </td>
	        <td width="60"> 
	          <div align="center" class="id"><%= oneboard.FBoardItem(i).FCheckFlag %></div>
	        </td>
	      </tr>
	    </table>
	  </div>
	</td>
</tr>
<%
next
%>
<tr>
	<td colspan="9" align="center" class="a">
		<% if oneboard.HasPreScroll then %>
		<a href="?page=<%= oneboard.StarScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oneboard.StarScrollPage to oneboard.ScrollCount + oneboard.StarScrollPage - 1 %>
			<% if i>oneboard.Totalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oneboard.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
	<!--
	<td>
		<table width="560" border="0" cellpadding="0" cellspacing="0" class="a" align="center">
		<tr valign="top"> 
		  <td> 
		    <div align="center"><span class="coment"><a href="#">◀</a></span><a href="#">[1]</a>2<a href="#">[3]</a><a href="#">[4]</a><a href="#">[5]</a><a href="#">[6]</a><a href="#">[7]</a><a href="#">[8]</a><a href="#">[9]</a><span class="coment"><a href="#">▶</a></span></div>
		  </td>
		</tr>
		</table>
	</td>
	-->
</tr>
</table>

<%
set oneboard = Nothing
%>
<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
