<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/lib/companybodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
dim site_name
dim id,name,gotopage
dim table_name

site_name = "tingmart"
table_name = "tbl_board_order"
id = request("id")
name = request("name")
gotopage = request("gotopage")
' count 증가
dim sqlput
sqlput = "update tbl_board_order set count=count+1 where id = "&id&" "
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

<form method="post" name="Form" action="doadmin_board_write.asp">

<%
    '게시물에 관한정보를 위한 select 문
    dim sql,mail,title,body,reg_date,count,thread,o_pos,o_depth,depth,pos,orderserial,mail_confirm
    sql = "SELECT id,name,mail,title,body,reg_date,count,thread,pos,depth "
    sql = sql + ",orderserial ,mail_confirm "
    sql = sql + " FROM tbl_board_order where id = '" + id + "' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            name = rsget("name")
            mail = rsget("mail")
            title = rsget("title")
            body = rsget("body")
            reg_date = rsget("reg_date")
            count = rsget("count")

            orderserial = rsget("orderserial")
            mail_confirm = rsget("mail_confirm")

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
    sql = "SELECT count(*) cnt from tbl_board_order where thread= '"+s_thread+"' and depth > '"+s_depth+"' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            plus_pos = rsget("cnt")
        else
            plus_pos = 0
        end if
    rsget.close
    pos = o_pos + plus_pos + 1

    ' 주문자의 신상명세를 조회하는 곳
    dim username, juminno, userphone, usercell, zipcode, useraddr, usermail, regdate, userpreaddr,birthday
    sql = "select top 1 *,(c.addr010_si + ' ' + c.addr010_gu) as userpreaddr from tbl_user_n a, tbl_logindata b, addr010tl c"
    sql = sql + " where (a.userid = b.userid) and (a.userid = '" + name + "') and c.addr010_zip1 = Left(a.zipcode,3) and c.addr010_zip2 = Right(a.zipcode,3)"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        username = rsget("username")
        juminno = rsget("juminno")
        userphone = rsget("userphone")
        usercell = rsget("usercell")
        zipcode = rsget("zipcode")
        userpreaddr = rsget("userpreaddr")
        useraddr = rsget("useraddr")
        birthday = rsget("birthday")
    end if
    rsget.close

    ' 총주문건수/총주문금액
    dim totCnt, totSum, avePrice
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from tbl_order_master where ipkumdiv not in ('0','1') and cancelyn = 'N'"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        totCnt = rsget("totCnt")
        totSum = rsget("totSum")
        if totCnt <> 0 then
            avePrice = totSum/totCnt
        end if
    end if
    rsget.close

    ' '유저별주문건수/총주문금액
    dim usrTotCnt, usrTotSum, usrAvePrice
    usrTotSum = 0
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from tbl_order_master where userid = '" + name + "' and ipkumdiv not in ('0','1') and cancelyn = 'N'"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        usrTotCnt = rsget("totCnt")
        usrTotSum = rsget("totSum")
        if usrTotCnt <> 0 then
            usrAvePrice = usrTotSum/usrTotCnt
        end if
    end if
    rsget.close

    ' '유저별 2개월내 주문건수/주문금액
'    dim fTotCnt, fTotSum, fAvePrice
'    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from tbl_order_master where userid = '" + name + "' and ipkumdiv not in ('0','1') and cancelyn = 'N' and datediff(day,regdate,getdate()) < 60"
'    rsget.Open sql,dbget,1
'    if  not rsget.EOF  then
'        fTotCnt = rsget("totCnt")
'        fTotSum = rsget("totSum")
'        if fTotCnt <> 0 then
'            fAvePrice = fTotSum/fTotCnt
'        end if
'    end if
'    rsget.close

    ' 유저별 결재완료후 취소건수/금액합계
    dim cTotCnt, cTotSum
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from   tbl_order_master where userid = '" + name + "' and ipkumdiv >= '5' and cancelyn = 'Y'"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        cTotCnt = rsget("totCnt")
        cTotSum = rsget("totSum")
    end if
    rsget.close

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
      주문번호 : <%=orderserial%> |
      글쓴이: <span class="id"><%=name%></span>| 날짜: <%=reg_date%></td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td class="a" height="5">
        유저명: <%=username%> | 생일 : <%=birthday%> |
        전화: <%=userphone%> | 핸드폰: <%=usercell%>
       </td>
    </tr>
    <tr>
      <td class="a" height="5">
        주소 : [<%=zipcode%>] <%=userpreaddr%> <%=useraddr%>
       </td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>

  </table>
  <table width="580" border="0" cellpadding="3" cellspacing="1">
    <tr>
      <td width="35" valign="top">
        <div align="right" class="a">내용 : </div>
      </td>
      <td width="506">
          <div align="left" class="a"><%=body%></div>
          <br>
        <input type="hidden" name="id" value="<%=id%>">
        <input type="hidden" name="pos" value="<%=pos%>">
        <input type="hidden" name="thread" value="<%=thread%>">
        <input type="hidden" name="depth" value="<%=depth%>">
        <input type="hidden" name="name" value="<%=name%>">
        <input type="hidden" name="table_name" value="<%=table_name%>">
        <input type="hidden" name="site_name" value="<%=site_name%>">
        </p>
        <p class="a"><br>
        </p>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="eeeeee">
          <tr>
            <td width="220">
              <div align="left" class="a">이메일 : <%=mail%></div>
            </td>
            <td width="92" class="a">
              <div align="right">이메일 수신여부 </div>
            </td>
            <td width="95" class="a">
              <div align="left">
<%          if mail_confirm = "Y" then %>
              수신요청함
<%          else %>
              수신요청안함
<%          end if %>
             </div>
              <input type="hidden" name="mail_confirm" value="<%=mail_confirm%>">
            </td>
            <td width="1">
              &nbsp;
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="55" align="left">
        <div align="left" class="a"><font color="#CCCCCC" class="a">답변제목: </font></div>
      </td>
      <td width="506">
        <input type="text" name="title" size="56" value="Re:<%=title%>">
      </td>
      </tr>
      <tr>
      <td width="55" align="left" valign="top">
        <div align="left" class="a"><font color="#CCCCCC" class="a">답변본문: </font></div>
      </td>
      <td width="506" valign="top">
        <textarea name="body" cols="56" rows="13"></textarea>
        <br>
        <p class="a"> </p>
<% if mail_confirm = "Y" then %>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="eeeeee">
          <tr>
            <td width="220">
              <div align="left" class="a">이메일 : <input type="text" name="mail" value="<%=mail%>"></div>
            </td>
            <td width="92" class="a">
              <div align="right">이메일 발송여부 </div>
            </td>
            <td width="48" class="a">
              <div align="left">
                <input type="radio" name="send_mail" value="Y" checked>
                Yes</div>
            </td>
            <td width="48">
              <div align="left">
                <input type="radio" name="send_mail" value="N">
                <span class="a">No</span></div>
            </td>
          </tr>
        </table>
<% end if %>
      </td>
    </tr>
  </table>
  <table width="81" border="0" align="center" cellpadding="5" cellspacing="0" height="15">
    <tr>
      <td>
        <div align="center"> <input type="image" src="/admin/images/reply_butten.gif" width="55" height="17"></div>
      </td>
      <td>
        <a href="javascript:history.back()"><img src="/admin/images/cancle_butten.gif" width="55" height="17" border="0"></a>
      </td>
      <td nowrap>
        <a href="docheckflag.asp?s_thread=<%=s_thread%>">처리완료</a>
      </td>
    </tr>
    <tr><td>&nbsp;</td></tr>
  </table>
</div>
</form>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->