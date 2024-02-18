<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "고객 게시판조회"
%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->
<%
dim site_name
site_name = request("site_name")
dim table_name
table_name = request("table_name")
dim gotopage
gotopage = request("gotopage")
%>

      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
      <form name=Form1 method="get" action="admin_board_list_all.asp">
        <tr> 
          <td> 
            <table width="350" border="0" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="100"> 
                  <div align="right"> 
                    <select name="site_name">
                      <option value="" >사이트선택</option>
                      <option value="" >----------</option>
                      <option value="" >전체</option>
                      <option value="10x10" <%if site_name="10x10" then response.write " selected"%>>10X10</option>
                      <option value="uto" <%if site_name="uto" then response.write " selected"%>>uto</option>
                      <option value="ugiljun" <%if site_name="ugiljun" then response.write " selected"%>>ugiljun</option>
                      <option value="yahoo" <%if site_name="yahoo" then response.write " selected"%>>yahoo</option>
                    </select>
                  </div>
                </td>
                <td width="10"> 
                  <select name="table_name">
                    <option value="" selected>게시판 선택 </option>
                    <option value="" >----------</option>
                    <option value="tbl_board_order" <%if table_name="tbl_board_order" then response.write " selected"%>>주문 배송</option>
                    <option value="tbl_board_site" <%if table_name="tbl_board_site" then response.write " selected"%>>사이트관련</option>
                  </select>
                  <input type="hidden" name="gotopage" value="<%=gotopage%>">
                </td>
                <td width="257"><input type="image" src="/admin/images/search2.gif" width="74" height="22"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
</form>
<br>

<% if table_name <> "tbl_board_site" then %>
<!-- 테이블이 명시 되었을경우임 -->

<table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td class="a"><img src="/admin/images/10x10_order_title.gif" width="124" height="31"></td>
  </tr>
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
  if site_name <> "" then
      query1 = query1 + " where site_name = '"+site_name+"'"
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
  SQL = SQL & "   "
  if site_name <> "" then
    SQL = SQL & " WHERE site_name = '"&site_name&"' "
  end if
  SQL = SQL & " ORDER BY thread desc, depth ) "
  if site_name <> "" then
    SQL = SQL & "  AND site_name = '"&site_name&"'  "
  end if
  SQL = SQL & " ORDER BY thread desc, depth "

dim inx,rownum,title,reg_date,isNew
dim s_id,s_name,s_site_name
dim i_depth ,i
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
                  <td width="80"><a href="" onclick="show_order_item('<%=rsget("orderserial")%>'); return false;"><%=rsget("orderserial")%></a></td>
                  <td>
<%

                   i_depth = CInt(rsget("depth"))
                   if i_depth > 1 then
                       for i=2 to i_depth 
                           response.write("&nbsp;&nbsp;")
                       next
                    response.write("<img src='/admin/images/re.gif' width='23' height='16'> ")   
                    end if
%>             
                   
                   <a href="admin_board_read.asp?id=<%=rsget("id")%>&table_name=tbl_board_order&site_name=<%=rsget("site_name")%>&name=<%=rsget("name")%>&gotopage=<%=gotopage%>">
                    <% if (rsget("check_flag")="D") then %>
                    <strike><%=title%></strike>(삭제됨)
                    <% else %>
                    <%=title%>
                    <% end if %>
                   </a>
                  </td>
                  <td width="90"> 
                    <div align="center">
                    
                    <%=rsget("name")%>
                    
                    </div>
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
      <% call gotoPageHTML3(gotopage, pagecount,"tbl_board_order",site_name)%>
      </span></div>
    </td>
  </tr>
</table>
<br>
<br>

<% end if %>

<% if table_name <> "tbl_board_order" then %>
<!-- 테이블이 명시 되었을경우임 -->

<table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td class="a"><img src="/admin/images/10x10_site_title.gif" width="124" height="31"></td>
  </tr>
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
  gotopage = request("gotopage")
  if gotopage = "" then gotopage = 1

  scale = 10
  page_scale = 10

  query1 = "select count(id) cnt from tbl_board_site "
  if site_name <> "" then
      query1 = query1 + " where site_name = '"+site_name+"'"
  end if
  rsget.Open query1,dbget,1
  recordcount = CInt(rsget("cnt"))
  rsget.close
  
  pagecount = int((recordcount-1)/scale) +1

        
  SQL = "SELECT TOP " & scale & " * "
  SQL = SQL & " FROM tbl_board_site "
  SQL = SQL & " WHERE id not in "
  SQL = SQL & "  (SELECT TOP " & ((gotopage - 1) * scale) & " id "
  SQL = SQL & "   FROM tbl_board_site "
  if site_name <> "" then
    SQL = SQL & "   where site_name = '"&site_name&"' "
  end if
  SQL = SQL & " ORDER BY thread desc, depth ) "
  if site_name <> "" then
    SQL = SQL & "  AND site_name = '"&site_name&"'  "
  end if
  SQL = SQL & " order by thread desc, depth "

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
                  <td> 
<%
                   i_depth = CInt(rsget("depth"))
                   if i_depth > 1 then
                       for i=2 to i_depth 
                           response.write("&nbsp;&nbsp;")
                       next
                    response.write("<img src='/admin/images/re.gif' width='23' height='16'> ")   
                    end if
%>                           
                   <a href="admin_board_read.asp?id=<%=rsget("id")%>&table_name=tbl_board_site&site_name=<%=rsget("site_name")%>&name=<%=rsget("name")%>&gotopage=<%=gotopage%>">
                   <% if (rsget("check_flag")="D") then %>
                    <strike><%=title%></strike>(삭제됨)
                   <% else %>
                    <%=title%>
                   <% end if %>
                   </a>
                  </td>
                  <td width="90"> 
                    <div align="center">
                    <%=rsget("name")%>
                    </div>
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
      <% call gotoPageHTML3(gotopage, pagecount,"tbl_board_site",site_name)%>
      </span></div>
    </td>
  </tr>
  <tr><td>&nbsp;</td></tr>
</table>

<% end if %>

<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
