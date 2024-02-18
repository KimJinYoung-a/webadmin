<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim site_name
site_name = "tingmart"
if (site_name = "") then
            response.write("<script>window.alert('Site 구분자가 넘어오지 않았습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if
dim table_name
table_name = "tbl_board_order"
if (table_name = "") then
            response.write("<script>window.alert('table 구분자가 넘어오지 않았습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if

dim id,name,gotopage
id = request("id")
name = request("name")
gotopage = request("gotopage")
' count 증가
dim sqlput
sqlput = "update "+table_name+" set count=count+1 where id = "&id&" "
rsput.Open sqlput,dbput,1

' 이전 다음버튼을 위해 id를 찾는부분
dim query1,before_id,after_id

%>
<script language="javascript">
<!--
function delMessage( table_name,site_name,messageId ){
    if( confirm( '이글을 삭제하시겠습니까?' ) ){
        URL = 'doboard_delete.asp?table_name='+table_name+'&site_name='+site_name+'&id='+messageId;
        document.location = URL;
    }
}
//-->
</script>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
      <form name=Form1 method="post" action="admin_board_list.asp">
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
<%
    '게시물에 관한정보를 위한 select 문
    dim sql,mail,title,body,reg_date,count,thread,o_pos,o_depth,depth,pos,orderserial,mail_confirm
    sql = "SELECT id,name,mail,title,body,reg_date,count,thread,pos,depth "
    if table_name = "tbl_board_order" then
      sql = sql + ",orderserial ,mail_confirm "
    end if
    sql = sql + " FROM "+table_name+" where id = '" + id + "' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then

            if CInt(rsget("depth")) > 1 then
                name = "10x10"
            else
                name = rsget("name")
            end if
            mail = rsget("mail")
            title = rsget("title")
            body = rsget("body")
            reg_date = rsget("reg_date")
            count = rsget("count")

        if table_name = "tbl_board_order" then
            orderserial = rsget("orderserial")
            mail_confirm = rsget("mail_confirm")
        end if

            title = db2html(title)
            body = db2html(body)
            mail = db2html(mail)

            body = Replace(body, vbcrlf, "<br>")
            thread = CInt(rsget("thread"))
            o_pos = CInt(rsget("pos"))
            o_depth = CInt(rsget("depth"))
            depth = depth + o_depth + 1

            reg_date = left(rsget("reg_date"),18)
        else
            response.write("<script>window.alert('해당자료가 존재하지 않습니다.');</script>")
            response.write("<script>history.back();</script>")
            dbget.close()	:	response.End
        end if
    rsget.close

    ' 자료의 위치를 찾는곳
    dim plus_pos,s_thread,s_depth
    s_thread = CStr(thread)
    s_depth = CStr(o_depth)
    sql = "SELECT count(*) cnt from "+table_name+" where thread= '"+s_thread+"' and depth > '"+s_depth+"' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            plus_pos = rsget("cnt")
        else
            plus_pos = 0
        end if
    rsget.close
    pos = o_pos + plus_pos + 1

%>

<div align="center"><br>
  <table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
    <tr>
      <td background="/admin/images/topbar_bg.gif" height="25" valign="middle">
        <div align="left">
          <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
            <tr>
              <td>
                <div align="left"><span class="a"><b>☞ <%=title%></b></span></div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr>
      <td class="a" height="5"> 사이트: <span class="id"><%=site_name%></span> |
<%  if table_name = "tbl_board_order" then %>
      주문번호 : <a href="" onclick="show_order_item('<%=orderserial%>'); return false;"><%=orderserial%></a> |
<%  end if %>
      글쓴이: <span class="id"><%=name%></span>| 날짜: <%=reg_date%></td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
  <table width="580" border="0" cellpadding="3" cellspacing="1">
    <tr>
      <td width="35" valign="top">
        <div align="right" class="a">내용 : </div>
      </td>
      <td width="506">
          <div class="a"><%=body%></div>
          <br>
      </td>
    </tr>
  </table>
  <table width="580" border="0" cellpadding="3" cellspacing="1">
    <tr>
      <td width="35" valign="top">
        &nbsp;
      </td>
      <td width="506">
        <div align="right">
      <% if o_depth = 1 then %>
        <a href="admin_board_write.asp?id=<%=id%>&table_name=<%=table_name%>&site_name=<%=site_name%>&name=<%=name%>&gotopage=<%=gotopage%>">
        <img src="/admin/images/reply_butten.gif" width="55" height="17" border="0">
        </a>
      <% end if %>
        </div>
      </td>
    </tr>
  </table>


<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->