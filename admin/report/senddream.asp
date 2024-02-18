<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "센드드림"
%>

<%
const Maxlines = 10
dim totalpage, totalnum, q

dim gotopage,fY,fM,fD,tY,tM,tD
dim fromDate,toDate,jnx,tmpStr,siteId,settle
dim showtype, IsAdmin


siteId = session("ssBctId")
if (siteID="10x10") then IsAdmin = true

dim searchId,mxlen
searchId = request("searchId")
showtype = request("showtype")
gotopage = request("gotopage")
settle = request("settle")
mxlen = request("mxlen")

fY = request("fY")
fM = request("fM")
fD = request("fD")

tY = request("tY")
tM = request("tM")
tD = request("tD")

''서동팔 수정..
''기본값적용..
If gotopage <> "" then
   session("gotopage") = CInt(gotopage)
else
   Session("gotopage") = 1
   gotopage = session("gotopage")
end if

gotopage= Cint(gotopage)

if (settle="") then settle ="D"

if (Not IsNumeric(mxlen)) or mxlen="" then mxlen =20


if showtype="2" then
	if (fY="") then fY = cstr(year(now()))
	if (fM="") then fM = cstr(month(now()))

	fromDate = DateSerial(fY, fM, 1)
	toDate = DateSerial(fY, fM+1, 1)
else
	if (fY="") then fY = cstr(year(now()))
	if (fM="") then fM = cstr(month(now()))
	if (fD="") then fD = cstr(day(now()))
	if (tY="") then tY = cstr(year(now()))
	if (tM="") then tM = cstr(month(now()))
	if (tD="") then tD = cstr(day(now()))

	fromDate = DateSerial(fY, fM, fD)
	toDate = DateSerial(tY, tM, tD+1)
end if

%>
      <table width="100%" border="0" cellpadding="0" cellspacing="3" bgcolor="#CCCCCC">
      <form name="bari" method="get" action="">
    	<input type="hidden" name="goTopage">
    	<input type="hidden" name="showtype" value="<%= showtype %>">
    	<input type="hidden" name="Xler" value="">

        <tr>
          <td width="5%"></td>
          <td width="600" class="a">
                  <select name="fY">
                    <!-- <option value="" <%if fY="" then response.write " selected"%>>년</option> -->
               <% for jnx=1 to 6 %>
               <%   tmpStr = "200"+CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fY=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="fM">
                    <!-- <option value="" <%if fM="" then response.write " selected"%>>월</option> -->
               <% for jnx=1 to 12 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fM=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="fD">
                    <!-- <option value="" <%if fD="" then response.write " selected"%>>일</option> -->
               <% for jnx=1 to 31 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if fD=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <span class="a">~</span>
                  <select name="tY">
                    <!-- <option value="" <%if tY="" then response.write " selected"%>>년</option> -->
               <% for jnx=1 to 6 %>
               <%   tmpStr = "200"+CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tY=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="tM">
                    <!--  <option value="" <%if tM="" then response.write " selected"%>>월</option> -->
               <% for jnx=1 to 12 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tM=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>
                  <select name="tD">
                    <!--  <option value="" <%if tD="" then response.write " selected"%>>일</option> -->
               <% for jnx=1 to 31 %>
               <%   tmpStr = CStr(jnx) %>
                    <option value="<%=tmpStr%>" <%if tD=tmpStr then response.write " selected"%>><%=tmpStr%></option>
               <% next %>
                  </select>

          </td>
          <td align="left" class="a">

          </td>
          <td align="right" class="a">
          	<input type="image" src="/images/search2.gif" width="74" height="22" align="middle">
          </td>
          <td width="5%"></td>
        </tr>
        <tr>
          <td width="5%"></td>
          <td width="600" class="a" valign="center">

          </td>
          <td >
          </td>
          <td class="a">
          </td>
          <td width="5%"></td>
        </tr>
      </table>
    </form>

<!-- main logic -->
<%
    '결과물에 대한 통계
    dim sql,tmp,cnt,total,wintotal

	cnt =0
	total = 0

	if settle = "M" then
        tmp = "total"
    elseif settle = "D" then
        tmp = "cnt"
    end if

    sql = "select count(id) as cnt"
	sql = sql & " from [db_senddream].[dbo].tbl_new_senddream2"
	sql = sql & " where regdate > '" + CStr(fromDate) + "'"
	sql = sql & " and regdate < '" + CStr(toDate) + "'"
	''sql = sql & " and ipkumdiv>3"
	''sql = sql & " group by userid"
	''sql = sql & " order by " + tmp + " desc"

	''response.write sql
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        total = rsget("cnt")
    end if
    rsget.Close

    sql = "select count(id) as cnt"
	sql = sql & " from [db_senddream].[dbo].tbl_new_senddream2"
	sql = sql & " where regdate > '" + CStr(fromDate) + "'"
	sql = sql & " and regdate < '" + CStr(toDate) + "'"
	sql = sql & " and iswinner='Y'"

    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        wintotal = rsget("cnt")
    end if
    rsget.Close
%>

      <table width="750" border="0" cellspacing="1" cellpadding="3" bgcolor="#EFBE00">
        <tr align="center">
          <td width="120" class="a"><font color="#FFFFFF">구분</font></td>
          <td class="a" width="780"><font color="#FFFFFF">내용</font></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td width="120" class="a">검색기간</td>

          	<td class="a" width="780"><%=fromDate%> ~ <%= DateSerial(tY, tM, tD) %></td>

        </tr>

        <tr bgcolor="#FFFFFF">
          <td width="120" class="a">총검색건수</td>
          <td class="a" width="780">총 <%= FormatNumber(FormatCurrency(total),0)%> 명</td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td width="120" class="a">총당첨자수</td>
          <td class="a" width="780">총 <%= FormatNumber(FormatCurrency(wintotal),0)%> 명</td>
        </tr>

      </table>
      <br>
<%
	dim maxcnt,sql2,n
	dim Cweek
	Cweek= array("일","월","화","수","목","금","토")

	sql = "select s.id, s.userid, s.sendname, s.reqname ,convert(varchar(100),s.mailcontents) as mailcontents,s.iswinner , i.imgsmall, i.itemid, T.userid as win "
	sql = sql & " from [db_senddream].[dbo].tbl_new_senddream2 s "
	sql = sql & " left join (select userid from [db_senddream].[dbo].tbl_new_senddream2 where iswinner='Y') as T on T.userid=s.userid"
	sql = sql & " , [db_item].[dbo].tbl_item_image i"
	sql = sql & " where regdate > '" + CStr(fromDate) + "'"
	sql = sql & " and regdate < '" + CStr(toDate) + "'"
	sql = sql & " and s.itemid=i.itemid"
	sql = sql & " order by s.id"
	rsget.Open sql,dbget,1

	maxcnt = 0

%>


	  <table width="750" border="0" cellspacing="1" cellpadding="3" >
	  <Form name="frm_senddream" method="post" OnSubmit="return SendDreamWinner(this);">
	  당첨 년월 (ex : 2003-09) <input type="text" name="yyyymm" value="" maxlength="7" size="7">
      <tr>
	   <td align="right"><input type="submit" value="선택한분 당첨!"></td>
	  </tr>
	  </table>

      <table width="90%" border="0" cellspacing="1" cellpadding="3" bgcolor="#6CA0B2" class="verdana_9pt">
        <tr align="center">
          <td width="50" align="center" ><font color="#FFFFFF">선택</font></td>
          <td width="50" align="center" ><font color="#FFFFFF">상품</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">아이디</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">이름</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">받는분</font></td>
          <td width="500"><font color="#FFFFFF">내용</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">당첨</font></td>
          <td width="60" align="center" ><font color="#FFFFFF">기존당첨</font></td>
        </tr>
<%
    if  not rsget.EOF  then
        dim pathname,fso,ofile,inx,ipkumdiv_name,subtotalprice_F,fileName,tmpFileName

        do until (rsget.EOF)
%>
        <tr bgcolor="#FFFFFF">
          <td width="50" align="center" ><input type="checkbox" name="cbox" value="<%= rsget("id") %>"></td>
          <td width="50" align="center" ><img src="http://image.10x10.co.kr/image/small/<%= GetImageSubFolderByItemid(rsget("itemid")) %>/<%= rsget("imgsmall") %>" ></td>
          <td width="60" align="center" valign="middle"><%= rsget("userid") %></td>
          <td width="60" align="center" valign="middle"><%= rsget("sendname") %></td>
          <td width="60" align="center" valign="middle"><%= rsget("reqname") %></td>
          <td width="500" align="left" ><%= db2html(rsget("mailcontents")) %></td>

          <% if rsget("iswinner")="Y" then %>
          <td width="60" align="center" ><font color="#FF0000">Y</font></td>
          <% else %>
          <td width="60" align="center" >N</td>
          <% end if %>
          <% if IsNull(rsget("win")) then %>
          <td width="60" align="center" ></td>
          <% else %>
          <td width="60" align="center" ><font color="#FF0000">Y</font></td>
          <% end if %>
        </tr>
<%
            rsget.MoveNext
        loop
%>
      <tr>
      </form>
      </table>

<%
rsget.close
%>

      <br>
      <div align="center"><br>
      </div>
<%  else %>

  <tr bgcolor="#FFFFFF">
    <td colspan="8">
      <center>해당기간에 거래내역이 없습니다.</center>
    </td>
  </tr>

<%  end if %>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
