<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/common/lib/pop_slide/classes/slidemanageCls.asp"-->
<%
'###########################################################
' Description :  �����̵� ���� list
' History : 2019-02-19 ����ȭ ����
'###########################################################
dim isusing , prevDate , i
dim page , mastercode , detailcode , menu , device
dim oSlideManage

menu = request("menu")
mastercode = request("mastercode")
detailcode = request("detailcode")
prevDate = request("prevDate")
isusing = request("isusing")
page = request("page")
device = request("device")

if page = "" then page = 1
if menu = "" then 
    response.write "<script>alter('�޴� ������ �ʿ��մϴ�.');self.close();</script>"
    response.end
end if 

set oSlideManage = new SlideListCls
    oSlideManage.FPageSize = 10
	oSlideManage.FCurrPage = page
	oSlideManage.FrectMasterCode = mastercode
	oSlideManage.FrectDetailCode = detailcode
    oSlideManage.FRectSelDate = prevDate
    oSlideManage.FRectIsusing = isusing
    oSlideManage.FRectMenu    = menu
    oSlideManage.FRectDevice    = device
	oSlideManage.getSlideList()
    
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript" src="/admin/common/lib/js/front.js"></script>
<script type="text/javascript">
    // slide ���
    function fnAddPopSlideManage(idx,m,d,device){
        var popwin = window.open('/admin/common/lib/pop_slide/pop_slide_manage_insert.asp?idx='+idx+'&menu=<%=menu%>&mastercode='+m+'&detailcode='+d+'&device='+device,'mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
        popwin.focus();
    }

    // ������ �̵�
	function NextPage(ipage)
	{
		document.frm.page.value= ipage;
		document.frm.submit();
	}

    function fnchglist() {
        document.frm.submit();
    }
</script>
</head>
<body>
<div class="contSectFix scrl">
    <div class="contHead">
		<div class="locate"><h2>��ȹ�� �����̵� ����</h2></div>
	</div>
	<div class="pad10">
        <div>
            <form name="frm" method="get" action="">
            <input type="hidden" name="page" value=""/>
            <input type="hidden" name="menupos" value="<%= request("menupos") %>"/>
            <input type="hidden" name="menu" value="<%=menu%>"/>
                <table class="tbType1 listTb">
                    <tr>
                        <td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
                        <td style="text-align:left;">
                            ��ȹ�� : <%=DrawSelectAllView("mastercode",mastercode,"fnchglist",menu)%>
                            <% if mastercode <> "" then %>
                                <%=DrawSelectDetailView("detailcode",mastercode,detailcode,"fnchglist",menu)%>
                            <% end if %>
                            ��뱸�� : 
                            <select name="isusing" class="select">
                                <option value="" <% if isusing="" then response.write "selected" %>>��ü
                                <option value="1" <% if isusing="1" then response.write "selected" %>>�����
                                <option value="0" <% if isusing="0" then response.write "selected" %>>������
                            </select>
                            ä�� : 
                            <select name="device" class="select">
                                <option value="" <% if device = "" then response.write "selected"%>>��ü
                                <option value="P" <% if device = "P" then response.write "selected"%>>PC
                                <option value="M" <% if device = "M" then response.write "selected"%>>M/A
                            </select>
                            &nbsp;&nbsp;
                            �������� :  <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" style="vertical-align:middle;"/>
                            <script language="javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "prevDate", trigger    : "prevDate_trigger",
                                    onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
                                });
                            </script>
                        </td>
                        <td width="50" bgcolor="<%= adminColor("gray") %>">
                            <input type="submit" class="button_s" value="�˻�">
                        </td>
                    </tr>
                </table>                
            </form>
        </div>
		<div class="tPad15">
            <table class="tbType1 listTb">
                <tr height="25" bgcolor="FFFFFF">
                    <td style="text-align:left;" colspan="8">
                        <div style="float:left;">
                            �˻���� : <b><%=oSlideManage.FtotalCount%></b>&nbsp;������ : <b><%= page %> / <%=oSlideManage.FtotalPage%></b>
                            <br/>
                            <span style="color:#ff0000"><strong>�� �̺�Ʈ ���� �������� ���� �̺�Ʈ �ڵ尡 ���� ��츸 ������ �˴ϴ�. ��</strong></span>
                        </div>
                        <div style="float:right;vertial-align:bottom;">
                            <input type="button" value="�̸�����" class="button" onclick="popSlideView('<%=mastercode%>','<%=detailcode%>','<%=menu%>');">                            
                            <input type="button" value="�űԵ��" class="button" onclick="fnAddPopSlideManage('0','<%=mastercode%>','<%=detailcode%>','<%=device%>');">
                        <div>
                    </td>
                </tr>
                <tr bgcolor="<%= adminColor("tabletop") %>" height="25" >
                    <th width="50">idx</th>
                    <th>����<br/>����2<br/>(�������� �ؽ�Ʈ)<br/>�̺�Ʈ�ڵ�</th>
                    <th>ä��</th>
                    <th>�̹���/������</th>
                    <th>��� ���� ������</th>
                    <th>��� ���� ������</th>
                    <th>��뿩��</th>
                    <th>�켱����</th>
                </tr>
                <%
                    for i=0 to oSlideManage.FResultCount - 1
                %>
                <tr bgcolor="<%=chkiif((oSlideManage.FItemList(i).IsEndDateExpired) or (oSlideManage.FItemList(i).FIsusing="0"),"#DDDDDD","#FFFFFF")%>">
                    <td><a href="javascript:fnAddPopSlideManage('<%=oSlideManage.FItemList(i).Fidx%>','','','');"><%= oSlideManage.FItemList(i).Fidx%></a></td>
                    <td><%= oSlideManage.FItemList(i).Ftitlename%><br/><br/><%= oSlideManage.FItemList(i).Fsubtitlename%><br/><br/><strong>[<a href="http://www.10x10.co.kr/event/eventmain.asp?eventid=<%= oSlideManage.FItemList(i).Feventid %>" target="_blank"><%= oSlideManage.FItemList(i).Feventid %></a>]</strong></td>
                    <td><%= chkiif(oSlideManage.FItemList(i).Fdevice="P","PC","M/A")%></td>
                    <td>
                        <% if oSlideManage.FItemList(i).Fisvideo = 1 then %>
                        ������
                        <% else %>
                            <% if oSlideManage.FItemList(i).Fimageurl = "" then %>
                            �̹��� �̵��
                            <% else %>
                            <img src="<%= oSlideManage.FItemList(i).Fimageurl %>" width="75"/>
                            <% end if %>
                        <% end if %>
                    </td>
                    <td><%= oSlideManage.FItemList(i).Fstartdate%><br/><br/>�̺�Ʈ������ : <span style="color:#ff0000"><%= oSlideManage.FItemList(i).Fevt_startdate%></span></td>
                    <td><%= oSlideManage.FItemList(i).Fenddate%><br/><br/>�̺�Ʈ������ : <span style="color:#ff0000"><%= oSlideManage.FItemList(i).Fevt_enddate%></span></td>
                    <td><%= chkiif(oSlideManage.FItemList(i).Fisusing="1","Y","N") %></td>
                    <td><%= oSlideManage.FItemList(i).Fsorting%></td>
                </tr>
                <% 
                    next 
                %>
                <tr bgcolor="#FFFFFF">
                    <td colspan="8" align="center" height="30">
                    <% if oSlideManage.HasPreScroll then %>
                        <a href="javascript:NextPage('<%= oSlideManage.StartScrollPage-1 %>');">[pre]</a>
                    <% else %>
                        [pre]
                    <% end if %>

                    <% for i=0 + oSlideManage.StartScrollPage to oSlideManage.FScrollCount + oSlideManage.StartScrollPage - 1 %>
                        <% if i>oSlideManage.FTotalpage then Exit for %>
                        <% if CStr(page)=CStr(i) then %>
                        <font color="red">[<%= i %>]</font>
                        <% else %>
                        <a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
                        <% end if %>
                    <% next %>

                    <% if oSlideManage.HasNextScroll then %>
                        <a href="javascript:NextPage('<%= i %>');">[next]</a>
                    <% else %>
                        [next]
                    <% end if %>
                    </td>
                </tr>
            </table>
		</div>
	</div>
</div>
<%
    SET oSlideManage = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->