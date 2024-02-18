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
if (site_name = "") then
            response.write("<script>window.alert('Site 구분자가 넘어오지 않았습니다.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if
dim table_name
table_name = request("table_name")
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
''query1 = "select top 1 id from "+table_name+" "_
'        &" where id not in (select id from "+table_name+" where id >= "&id&" ) "_
'        &" order by id desc "
'rsget.Open query1,dbget,1

'if not rsget.EOF  then
'    before_id = rsget("id")
'end if
'rsget.Close
'query1 = "select top 1 id from "+table_name+" "_
'        &" where id > "&id&" "_
'        &" order by id "

'rsget.Open query1,dbget,1
'if not rsget.EOF  then
'    after_id = rsget("id")
'end if
'rsget.Close

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
<table width="580" border="0" align="center" cellpadding="0" cellspacing="3">
  <tr> 
    <td background="/admin/images/topbar_bg.gif" height="25" valign="middle"> 
      <div align="left"> 
        <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
          <tr> 
            <td> 
              <div align="left"><b>☞ 관련글</b></div>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
<%

      dim index
      s_thread = CStr(thread)
      sql = "SELECT id,title,name,reg_date,count,depth,site_name FROM "+table_name+" "
      sql = sql + " WHERE thread = '" + s_thread + "'  order by pos"
      rsget.Open sql,dbget,1
      if  not rsget.EOF  then
          rsget.Movefirst
          index = 0
          do until rsget.EOF  
%>  
  <tr> 
    <td height="5" valign="top"> 
      <div align="center"> 
        <table width="560" border="0" cellpadding="0" cellspacing="2">
          <tr> 
            <td> 
              <div align="center"> 
                <table width="560" border="0" cellpadding="0" cellspacing="0" class="a">
                  <tr  onMouseOver='this.style.backgroundColor="#eeeeee"' onClick="" onMouseOut='this.style.backgroundColor="#ffffff"'  bgcolor=white height=20> 
                    <td> 
<%
                 dim i_depth ,i, isNew
                 i_depth = CInt(rsget("depth"))
                 if i_depth > 1 then
                     for i=2 to i_depth 
                         response.write("&nbsp;")
                     next
                  response.write("<img src='/admin/images/re.gif' width='23' height='16'>")   
                  end if

                  title = db2html(rsget("title"))
                  if len(title) > 50 then
                      title = left(title,45)&"..."
                  end if
                  if DateDiff("d",CDate(rsget("reg_date")),Date) < 2 then
                      isNew = "T"
                  end if     
%>             
                      <span class="vadana">
                        <a href="admin_board_read.asp?id=<%=rsget("id")%>&table_name=<%=table_name%>&site_name=<%=rsget("site_name")%>&name=<%=rsget("name")%>&gotopage=<%=gotopage%>"><%= title %></a></span>
                        <span class="red"><% if isNew = "T" then response.write("new")%></span>                       
                    </td>
                    <td width="80"> 
                      <div align="center" class="id">
                      <% if i_depth > 1 then %>
                      10x10
                      <% else %>
                      <%=rsget("name")%>
                      <% end if %>
                      </div>
                    </td>
                    <td width="80" class="vadana"> 
                      <div align="center"><%=Left(rsget("reg_date"),11)%></div>
                    </td>
                    <td width="30" class="vadana"> 
                      <div align="center"><%=rsget("count")%></div>
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
          <tr> 
            <td> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td height="2"><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
<%      
         rsget.MoveNext
         index = index + 1
      loop
end if
rsget.close
%>   
    
  </table>  

</div>

<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
