<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim site_name
dim table_name
dim gotopage
dim research,check_flag

research = request("research")
check_flag = request("check_flag")
site_name="tingmart"

if research="" then check_flag="on"

%>
<table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
  <form name="frm" method="get" action="">
  <input type="hidden" name="research" value="on">
  <tr>
    <td class="a"><img src="/admin/images/10x10_order_title.gif" width="124" height="31"></td>
    <td width="120" class="a"><input type="checkbox" name="check_flag" <% if check_flag="on" then response.write "checked" %> >미처리만 보기</td>
    <td width="80"><a href="javascript:document.frm.submit()"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
  </tr>
  </form>
</table>
<table width="630" border="0" align="center">
  <tr>
    <td  valign="top">
      <table width="620" border="0" cellpadding="0" cellspacing="3" align="center">
        <tr>
          <td height="25" valign="middle" class="top_bg">
            <div align="center">
              <table width="610" border="0" cellpadding="0" cellspacing="0" class="a">
                <tr>
                  <td width="46">
                    <div align="left">사이트</div>
                  </td>
                  <td width="40">번호</td>
                  <td width="80">주문번호</td>
                  <td>제 목</td>
                  <td width="90">
                    <div align="center">글쓴이 </div>
                  </td>
                  <td width="94">
                    <div align="center">날 짜</div>
                  </td>
                </tr>
              </table>
            </div>
          </td>
        </tr>

<%
  dim scale,page_scale,total,query1
  Dim SQL
  Dim pagecount, recordcount

  gotopage = request("gotopage")
  if gotopage = "" then gotopage = 1

  scale = 10
  page_scale = 10

  query1 = "select count(id) cnt from tbl_board_order "
  query1 = query1 + " where id<>0"
  if check_flag="on" then
  	query1 = query1 + " and check_flag ='N' "
  end if

  if site_name <> "" then
      query1 = query1 + " and site_name = '"+site_name+"'"
  end if
  rsget.Open query1,dbget,1
  recordcount = CInt(rsget("cnt"))
  rsget.close

  pagecount = int((recordcount-1)/scale) +1


  SQL = "SELECT TOP " & scale & " * "
  SQL = SQL & " FROM tbl_board_order "
  SQL = SQL & " WHERE id not in "
  SQL = SQL & "  (SELECT TOP " & ((gotopage - 1) * scale) & " id "
  SQL = SQL & "   FROM tbl_board_order "
  SQL = SQL & "   WHERE  id<>0"
  if check_flag="on" then
  	SQL = SQL + " and check_flag ='N' "
  end if

  if site_name <> "" then
    SQL = SQL & " AND site_name = '"&site_name&"' "
  end if
  SQL = SQL & " ORDER BY thread desc, depth  ) "
  if site_name <> "" then
    SQL = SQL & "  AND site_name = '"&site_name&"'  "
  end if
  if check_flag="on" then
  	SQL = SQL + " and check_flag ='N' "
  end if

  SQL = SQL & " ORDER BY thread desc, depth  "

dim inx,rownum,title,reg_date,isNew
dim s_id,s_name,s_site_name
rownum = (gotopage-1)*10

rsget.Open SQL,dbget,1
if  not rsget.EOF  then
        rsget.Movefirst
        do until rsget.EOF

        title = db2html(rsget("title"))
        if len(title) > 50 then
            title = left(title,45)&"..."
        end if
        reg_date = Left(rsget("reg_date"),11)
        rownum = rownum + 1
        if DateDiff("d",CDate(rsget("reg_date")),Date) < 2 then
            isNew = "T"
        end if
%>

        <tr>
          <td>
            <div align="center">
              <table width="610" border="0" cellpadding="0" cellspacing="0" class="a">
                <TR  onmouseover='this.style.backgroundColor="#eeeeee"' onclick="" onmouseout='this.style.backgroundColor="#ffffff"' bgColor=white height=20>
                  <td width="46">
                    <div align="left"><%=rsget("site_name")%></div>
                  </td>
                  <td width="40"><%=rsget("id")%></td>
                  <td width="80"><%=rsget("orderserial")%></td>
                  <td>
                   <% if rsget("check_flag")="N" then %>
                   <a href="boardwrite.asp?id=<%=rsget("id")%>&gotopage=<%=gotopage%>">
                    <%=title%>
                   </a>
                   <% else %>
                   <a href="boardread.asp?id=<%=rsget("id")%>&gotopage=<%=gotopage%>">
                    <%=title%>
                   </a>
                   <% end if %>
                  </td>
                  <td width="90">
                    <div align="center"><%=rsget("name")%></div>
                  </td>
                  <td width="94">
                    <div align="center"><%=reg_date%></div>
                  </td>
                </tr>
              </table>
            </div>
          </td>
        </tr>

        <tr>
          <td>
            <div align="center"><img src="/admin/images/w_dot.gif" width="580" height="1"></div>
          </td>
        </tr>

<%
            rsget.MoveNext

        loop
end if
rsget.close
%>
      </table>
    </td>
  </tr>
</table>
<table width="560" border="0" cellpadding="0" cellspacing="0" class="a" align="center">
  <tr valign="top">
    <td>
      <div align="center"><span class="coment">
      <% call gotoPageHTML2(gotopage, pagecount,"tbl_board_order",site_name)%>
      </span></div>
    </td>
  </tr>
</table>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->