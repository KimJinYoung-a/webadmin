<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "�� �Խ��ǰ���"
%>
<%
dim site_name
site_name = request("site_name")
if (site_name = "") then
            response.write("<script>window.alert('Site �����ڰ� �Ѿ���� �ʾҽ��ϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if
dim table_name
table_name = request("table_name")
if (table_name = "") then
            response.write("<script>window.alert('table �����ڰ� �Ѿ���� �ʾҽ��ϴ�.');</script>")
            response.write("<script>history.go(-1);</script>")
            dbget.close()	:	response.End
end if

dim id,name,gotopage
id = request("id")
name = request("name")
gotopage = request("gotopage")
' count ����
dim sqlput
sqlput = "update [db_board].[10x10]."+table_name+" set count=count+1 where id = "&id&" "
rsput.Open sqlput,dbput,1

' ���� ������ư�� ���� id�� ã�ºκ�
dim query1,before_id,after_id
'query1 = "select top 1 id from "+table_name+" "_
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
    if( confirm( '�̱��� �����Ͻðڽ��ϱ�?' ) ){
        URL = 'doboard_delete.asp?table_name='+table_name+'&site_name='+site_name+'&id='+messageId;
        document.location = URL;
    }
}
//-->
</script>
<form name=Form1 method="post" action="admin_board_list.asp">
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
        <tr>
          <td>
            <table width="350" border="0" cellpadding="0" cellspacing="3">
              <tr>
                <td width="100">
                  <div align="right">
                    <select name="site_name">
                      <option value="" >����Ʈ����</option>
                      <option value="" >----------</option>
                      <option value="" >��ü</option>
                      <option value="10x10" <%if site_name="10x10" then response.write " selected"%>>10X10</option>
                      <option value="uto" <%if site_name="uto" then response.write " selected"%>>uto</option>
                      <option value="ugiljun" <%if site_name="ugiljun" then response.write " selected"%>>ugiljun</option>
                      <option value="yahoo" <%if site_name="yahoo" then response.write " selected"%>>yahoo</option>
                    </select>
                  </div>
                </td>
                <td width="10">
                  <select name="table_name">
                    <option value="" selected>�Խ��� ���� </option>
                    <option value="" >----------</option>
                    <option value="tbl_board_order" <%if table_name="tbl_board_order" then response.write " selected"%>>�ֹ� ���</option>
                    <option value="tbl_board_site" <%if table_name="tbl_board_site" then response.write " selected"%>>����Ʈ����</option>
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
<form method="post" name="Form" action="doadmin_board_write.asp">

<%
    '�Խù��� ���������� ���� select ��
    dim sql,mail,title,body,reg_date,count,thread,o_pos,o_depth,depth,pos,orderserial,mail_confirm
    sql = "SELECT id,name,mail,title,body,reg_date,count,thread,pos,depth "
    if table_name = "tbl_board_order" then
      sql = sql + ",orderserial ,mail_confirm "
    end if
    sql = sql + " FROM [db_board].[10x10]."+table_name+" where id = '" + id + "' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            name = rsget("name")
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
            response.write("<script>window.alert('�ش��ڷᰡ �������� �ʽ��ϴ�.');</script>")
            response.write("<script>history.back();</script>")
            dbget.close()	:	response.End
        end if
    rsget.close

    ' �ڷ��� ��ġ�� ã�°�
    dim plus_pos,s_thread,s_depth
    s_thread = CStr(thread)
    s_depth = CStr(o_depth)
    sql = "SELECT count(*) cnt from [db_board].[10x10]."+table_name+" where thread= '"+s_thread+"' and depth > '"+s_depth+"' "
    rsget.Open sql,dbget,1
        if  not rsget.EOF  then
            plus_pos = rsget("cnt")
        else
            plus_pos = 0
        end if
    rsget.close
    pos = o_pos + plus_pos + 1

    ' �ֹ����� �Ż���� ��ȸ�ϴ� ��
    dim username, juminno, userphone, usercell, zipcode, useraddr, usermail, regdate, userpreaddr,birthday
    sql = "select top 1 *,(c.addr030_si + ' ' + c.addr030_gu) as userpreaddr from [db_user].[10x10].tbl_user_n a, [db_user].[10x10].tbl_logindata b, [db_zipcode].[10x10].addr030tl c"
    sql = sql + " where (a.userid = b.userid) and (a.userid = '" + name + "') and c.addr030_zip1 = Left(a.zipcode,3) and c.addr030_zip2 = Right(a.zipcode,3)"
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

    ' ���ֹ��Ǽ�/���ֹ��ݾ�
    dim totCnt, totSum, avePrice
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from [db_order].[10x10].tbl_order_master where ipkumdiv not in ('0','1') and cancelyn = 'N'"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        totCnt = rsget("totCnt")
        totSum = rsget("totSum")
        if totCnt <> 0 then
            avePrice = totSum/totCnt
        end if
    end if
    rsget.close

    ' '�������ֹ��Ǽ�/���ֹ��ݾ�
    dim usrTotCnt, usrTotSum, usrAvePrice
    usrTotSum = 0
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from [db_order].[10x10].tbl_order_master where userid = '" + name + "' and ipkumdiv not in ('0','1') and cancelyn = 'N'"
    rsget.Open sql,dbget,1
    if  not rsget.EOF  then
        usrTotCnt = rsget("totCnt")
        usrTotSum = rsget("totSum")
        if usrTotCnt <> 0 then
            usrAvePrice = usrTotSum/usrTotCnt
        end if
    end if
    rsget.close

    ' '������ 2������ �ֹ��Ǽ�/�ֹ��ݾ�
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

    ' ������ ����Ϸ��� ��ҰǼ�/�ݾ��հ�
    dim cTotCnt, cTotSum
    sql = "select count(*) totCnt,isnull(sum(subtotalprice),0) totSum from  [db_order].[10x10].tbl_order_master where userid = '" + name + "' and ipkumdiv >= '5' and cancelyn = 'Y'"
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
                <div align="left"><span class="a"><b>�� <%=title%></b></span></div>
              </td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr>
      <td class="a" height="5"> ����Ʈ: <span class="id"><%=site_name%></span> |
<%  if table_name = "tbl_board_order" then %>
      �ֹ���ȣ : <a href="" onclick="show_order_item('<%=orderserial%>'); return false;"><%=orderserial%></a> |
<%  end if %>
      �۾���: <span class="id"><%=name%></span>| ��¥: <%=reg_date%></td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
<%  if table_name = "tbl_board_order" then %>
    <tr>
      <td class="a" height="5">
        ������: <%=username%> | ���� : <%=birthday%> |
        ��ȭ: <%=userphone%> | �ڵ���: <%=usercell%>
       </td>
    </tr>
    <tr>
      <td class="a" height="5">
        �ּ� : [<%=zipcode%>] <%=userpreaddr%> <%=useraddr%>
       </td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
    <tr>
      <td class="a" height="5">
<!--        10x10 �ֹ��Ǽ�: <%=totCnt%> �ݾ�:<%=FormatCurrency(totSum)%> ��ձ��ž�:<%=FormatCurrency(avePrice)%> | <br>     -->
        <%=name%> �ֹ��Ǽ�: <%=usrTotCnt%> �ݾ�:<%=FormatCurrency(usrTotSum)%> ��ձ��ž�:<%=FormatCurrency(usrAvePrice)%> | <br>
        <%=name%> �ֹ���ҰǼ�: <%=cTotCnt%> ��ұݾ�:<%=FormatCurrency(0)%> <br>
       </td>
    </tr>
    <tr>
      <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
    </tr>
  </table>
<%  end if %>
  <table width="580" border="0" cellpadding="3" cellspacing="1">
    <tr>
      <td width="35" valign="top">
        <div align="right" class="a">���� : </div>
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
<% if table_name = "tbl_board_order" then %>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="eeeeee">
          <tr>
            <td width="220">
              <div align="left" class="a">�̸��� : <%=mail%></div>
            </td>
            <td width="92" class="a">
              <div align="right">�̸��� ���ſ��� </div>
            </td>
            <td width="95" class="a">
              <div align="left">
<%          if mail_confirm = "Y" then %>
              ���ſ�û��
<%          else %>
              ���ſ�û����
<%          end if %>
             </div>
              <input type="hidden" name="mail_confirm" value="<%=mail_confirm%>">
            </td>
            <td width="1">
              &nbsp;
            </td>
          </tr>
        </table>
<% end if %>
      </td>
    </tr>
    <tr>
      <td width="55" align="left">
        <div align="left" class="a"><font color="#CCCCCC" class="a">�亯����: </font></div>
      </td>
      <td width="506">
        <input type="text" name="title" size="56" value="Re:<%=title%>">
      </td>
      </tr>
      <tr>
      <td width="55" align="left" valign="top">
        <div align="left" class="a"><font color="#CCCCCC" class="a">�亯����: </font></div>
      </td>
      <td width="506" valign="top">
        <textarea name="body" cols="56" rows="13"></textarea>
        <br>
        <p class="a"> </p>
<% if table_name = "tbl_board_order" and mail_confirm = "Y" then %>
        <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="eeeeee">
          <tr>
            <td width="220">
              <div align="left" class="a">�̸��� : <input type="text" name="mail" value="<%=mail%>"></div>
            </td>
            <td width="92" class="a">
              <div align="right">�̸��� �߼ۿ��� </div>
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
        <a href="docheckflag.asp?s_thread=<%=s_thread%>&table_name=<%=table_name%>">ó���Ϸ�</a>
      </td>
    </tr>
    <tr><td>&nbsp;</td></tr>
  </table>
</div>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
