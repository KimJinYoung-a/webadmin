<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "배송유의사항"

dim yyyy1,mm1,dd1
dim nowdate

nowdate = now

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

if (yyyy1="") then
	yyyy1 = Left(CStr(nowdate),4)
	mm1   = Mid(CStr(nowdate),6,2)
	dd1   = Mid(CStr(nowdate),9,2)
end if

%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->
<br>
<div align="center">
	<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
	<form name="frm" action="dodeliverywrite.asp" method="post">
          <tr> 
            <td class="a" width="409"><b><img src="/admin/images/mini_icon.gif" width="17" height="17"> 
              배송유의사항 쓰기</b></td>
            <td class="a"> 
              <div align="right"> </div>
            </td>
          </tr>
        </table>
        <br>
        <table width="750" border="0" cellpadding="3" cellspacing="1">
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="6"> 
              <div align="right"><font color="#CCCCCC" class="a">날짜 : </font></div>
            </td>
            <td width="407" height="-2"> 
              <select name="yyyy1">
                  	<option value="<%= yyyy1 %>" selected><%= yyyy1 %></option>
                    <option value="2001" >2001</option>
                  	<option value="2002" >2002</option>
                    <option value="2003" >2003</option>
                    <option value="2004" >2004</option>
                    <option value="2005" >2005</option>
                    <option value="2006" >2006</option>
                  </select>
                  <select name="mm1">
                  	<option value="<%= mm1 %>" selected><%= mm1 %></option>
                    <option value="01" >01</option>
                    <option value="02" >02</option>
                    <option value="03" >03</option>
                    <option value="04" >04</option>
                    <option value="05" >05</option>
                    <option value="06" >06</option>
                    <option value="07" >07</option>
                    <option value="08" >08</option>
                    <option value="09" >09</option>
                    <option value="10" >10</option>
                    <option value="11" >11</option>
                    <option value="12" >12</option>
                  </select>
                  <select name="dd1">
                  	<option value="<%= dd1 %>" selected><%= dd1 %></option>
                    <option value="01" >01</option>
                    <option value="02" >02</option>
                    <option value="03" >03</option>
                    <option value="04" >04</option>
                    <option value="05" >05</option>
                    <option value="06" >06</option>
                    <option value="07" >07</option>
                    <option value="08" >08</option>
                    <option value="09" >09</option>
                    <option value="10" >10</option>
                    <option value="11" >11</option>
                    <option value="12" >12</option>
                    <option value="13" >13</option>
                    <option value="14" >14</option>
                    <option value="15" >15</option>
                    <option value="16" >16</option>
                    <option value="17" >17</option>
                    <option value="18" >18</option>
                    <option value="19" >19</option>
                    <option value="20" >20</option>
                    <option value="21" >21</option>
                    <option value="22" >22</option>
                    <option value="23" >23</option>
                    <option value="24" >24</option>
                    <option value="25" >25</option>
                    <option value="26" >26</option>
                    <option value="27" >27</option>
                    <option value="28" >28</option>
                    <option value="29" >29</option>
                    <option value="30" >30</option>
                    <option value="31" >31</option>
                  </select>
              </td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="7"> 
              <div align="right"><font color="#CCCCCC" class="a">사이트 : </font></div>
            </td>
            <td width="407" height="-1"> 
              <% drawSelectBox "sitename","" %>
              <span class="a">사이트를 선택하세요</span></td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="2"> 
              <div align="right"><font color="#CCCCCC" class="a">고객명 : </font></div>
            </td>
            <td width="407" height="2"> 
              <input type="text" name="buyname" size="15">
            </td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="6"> 
              <div align="right"><font color="#CCCCCC" class="a">주문번호 : </font></div>
            </td>
            <td width="407" height="6"> 
              <input type="text" name="orderserial" size="20" maxlength="32">
              <span class="a">주문번호를 정확하게 입력하세요</span></td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="7"> 
              <div align="right"><font color="#CCCCCC" class="a">글쓴이 : </font></div>
            </td>
            <td width="407" height="7">
              <select name="writer">
                <option selected>선택</option>
                <option value="winnie" >최은희</option>
                <option value="moon" >이문재</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee" height="6"> 
              <div align="right"><font color="#CCCCCC" class="a">제목 : </font></div>
            </td>
            <td width="407" height="6"> 
              <input type="text" name="title" size="54" maxlength="128">
            </td>
          </tr>
          <tr> 
            <td width="100" bgcolor="#eeeeee"> 
              <div align="right" class="a"><font color="#CCCCCC" class="a">주의사항 
                내용 : </font></div>
            </td>
            <td> 
              <textarea name="txmemo" cols="53" rows="15"></textarea>
            </td>
          </tr>
        </table>
        <table border="0" align="center" cellpadding="0" cellspacing="5">
          <tr> 
            <td height="2"> 
              <div align="right"> 
                <p><a href="javascript:AnDeliveryWrite(frm)"><img src="/admin/images/write_butten.gif" width="55" border="0"></a></p>
              </div>
            </td>
            <td valign="top" height="2"> 
              <div align="center"><a href="javascript:history.back()"><img src="/admin/images/cancle_butten.gif" width="55" border="0"></a></div>
            </td>
          </tr>
        </table>
       </form> 
</div>
<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
