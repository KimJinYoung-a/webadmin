<%@ language=vbscript %>
<% option explicit %>
<% 
Response.AddHeader "Cache-Control","no-cache" 
Response.AddHeader "Expires","0" 
Response.AddHeader "Pragma","no-cache" 
%> 
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/bct_admin_header.asp"-->
<!-- #include virtual="/lib/classes/noreplyboardcls.asp"-->
<%
const MenuPos1 = "Admin"
const MenuPos2 = "배송유의사항"

dim i
dim boardid, oneboard,writer
boardid = request("id")
writer = request("writer")

if boardid="" then
	boardid =0
end if

set oneboard = new CNoReplyBoard
oneboard.readBoard boardid

%>
<!-- #include virtual="/admin/bct_admin_menupos.asp"-->

<br>
        <table width="630" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td class="a" width="409"><b><img src="/admin/images/mini_icon.gif" width="17" height="17"> 
              배송유의사항 읽기</b></td>
            <td class="a"> 
              <div align="right"> </div>
            </td>
          </tr>
        </table>
        <br>
        <table width="630" border="0" cellpadding="0" cellspacing="0" class="a" align="center">
          <tr> 
            <td> 
              <div align="right"> 
                <table width="102" border="0" cellpadding="0" cellspacing="5">
                  <tr> 
                    <td><a href="bct_admin_deliver.asp"><img src="/admin/images/list_butten.gif" width="55" height="17" border="0"></a></td>
                    <% if oneboard.FBoardItem(0).FPreID <>0 then %>
                    <td><a href="?id=<%=oneboard.FBoardItem(0).FPreID %>"><img src="/admin/images/pre.gif" height="17" border="0"></a></td>
                    <% else %>
                    <td></td>
                    <% end if %>
                    
                    <% if oneboard.FBoardItem(0).FNextID <>0 then %>
                    <td><a href="?id=<%=oneboard.FBoardItem(0).FNextID %>"><img src="/admin/images/next.gif" height="17" border="0"></a></td>
                    <% else %>
                    <td></td>
                    <% end if %>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
        </table>
        <table width="630" border="0" align="center" cellpadding="0" cellspacing="3">
          <tr> 
            <td background="/admin/images/topbar_bg.gif" height="25" valign="middle"> 
              <div align="left"> 
                <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
                  <tr> 
                    <td> 
                      <div align="left"><b>☞ <%= oneboard.FBoardItem(0).FTitle %> </b></div>
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
          <tr> 
            <td class="a" height="5"> 고객명 : <b><%= oneboard.FBoardItem(0).FBuyName %></b> |주문번호 : <b><%= oneboard.FBoardItem(0).FOrderSerial %> 
              </b>| 날짜: <b><%= oneboard.FBoardItem(0).FMatchDate %> </b>| 글쓴이 : <%= oneboard.FBoardItem(0).FWriter %><b> </b>| 처리여부 : <span class="id"><%= oneboard.FBoardItem(0).FCheckFlag %></span></td>
          </tr>
          <tr> 
            <td><img src="/admin/images/w_dot.gif" width="580" height="1"></td>
          </tr>
        </table>
        <!-- 본문 -실제글 읽는부분-->
        <table width="630" border="0" align="center" cellpadding="0" cellspacing="2">
          <tr> 
            <td class="a" valign="top">
            	<%= oneboard.FBoardItem(0).FMemo %>
            </td>
          </tr>
          <tr> 
            <td class="a" valign="top"> 
              <div align="right"> 
                <table width="350" border="0" cellspacing="3" cellpadding="0">
                  <form name="frm" method="post" action="doeditdelivery.asp" >
                  <input type="hidden" name="mode" value="">
                  <input type="hidden" name="id" value="<%= oneboard.FBoardItem(0).FID %>">
                  <tr> 
                    <td class="a" width="118"> 
                      <div align="right"><img src="/admin/images/yesorno.gif" width="95" height="28"> 
                      </div>
                    </td>
                    <td width="75"> 
                      <input type="radio" name="checkflag" value="N" <% if oneboard.FBoardItem(0).FCheckFlag="N" then response.write "checked" %>>
                      <span class="a">접수</span></td>
                    <td width="75"> 
                      <input type="radio" name="checkflag" value="Y" <% if oneboard.FBoardItem(0).FCheckFlag="Y" then response.write "checked" %>>
                      <span class="a">완료</span></td>
                    <td width="50" class="a"><b><a href="javascript:AnCheckDelivery(frm)"><img src="/admin/images/baesong_save.gif" width="63" height="28" border="0"></a></b></td>
                    <td width="40" class="a"><a href="javascript:AnDeleteDelivery(frm)"><img src="/admin/images/baesong_del.gif" width="63" height="28" border="0"></a></td>
                  </tr>
                  </form>
                </table>
              </div>
            </td>
          </tr>
        </table>
        <br>
        <!-- 메모하는곳-->
        <table width="630" border="0" align="center" cellpadding="0" cellspacing="5">
          <tr> 
            <td background="/admin/images/topbar_bg.gif" height="25" valign="middle"> 
              <div align="left"> 
                <table width="520" border="0" cellpadding="0" cellspacing="0" class="a">
                  <tr> 
                    <td> 
                      <div align="left"><b>☞ 추가사항</b></div>
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
          <% for i=0 to oneboard.FBoardItem(0).FcommentCount-1 %>
          <tr> 
            <td height="2" valign="top"> 
              <div align="left" class="a"> 
                <table border="0" cellspacing="0" cellpadding="0" width="630">
                  <tr> 
                    <td class="a" height="13" width="100"> 
                      <div align="center"><b><%= oneboard.FBoardItem(0).FComItem(i).FWriter %></b></div>
                    </td>
                    <td class="a" height="13"> 
                      <div align="left"><%= oneboard.FBoardItem(0).FComItem(i).FComment %></div>
                    </td>
                    <td height="13" width="120"> 
                      <div align="center" class="id"><%= Left(oneboard.FBoardItem(0).FComItem(i).FRegDate,16) %></div>
                    </td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
          <% next %>
        </table>
        <table width="630" border="0" align="center" cellpadding="0" cellspacing="5">
          <form name="frmcom" method="post" action="dodeliverycom.asp">
          <input type="hidden" name="masterid" value="<%=oneboard.FBoardItem(0).FID %>">
          <tr> 
            <td height="2" width="100" class="a"> 
              <div align="center"> 
                <p><img src="/admin/images/baesong_coment.gif"></p>
              </div>
            </td>
            <td valign="top" height="2" width="568"> 
              <div align="left">
                <input type="text" name="tx_com" size="53">
                <% drawSelectBoxWriter writer %>
                <a href="javascript:AnWriteDeliveryCom(frmcom)"><img src="/admin/images/write_butten.gif" width="55" border="0"></a></div>
            </td>
          </tr>
          </form>
        </table>

<%
set oneboard = Nothing
%>
<!-- #include virtual="/admin/bct_admin_tail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
