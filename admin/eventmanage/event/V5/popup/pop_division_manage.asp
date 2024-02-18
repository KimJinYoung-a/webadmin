<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/divisionCls.asp"-->
<%
'###############################################
' PageName : pop_division_manage.asp
' Discription : �̺�Ʈ ���� ���� �˾�
' History : 2019.02.21 ������
'###############################################

dim gidx, page , mode

gidx = request("gidx")
page = request("page")

if gidx="" then gidx=0
if page="" then page=1

if gidx = 0 then 
    mode = "gubunAdd"
else
    mode = "gubunModify"
end if 

dim oDivision,oDivisionList

set oDivision = new DivisionCls
oDivision.FrectGcode = gidx
oDivision.getOneGroupItem

set oDivisionList = new DivisionCls
oDivisionList.FPageSize=20
oDivisionList.FCurrPage= page
oDivisionList.getGroupList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function SavePosCode(frm){
    <% if mode = "gubunAdd" then %>
    if((!frm.gubuncode[0].selected)&&(!frm.gubuncode[1].selected)){
        alert('��ȹ�� ������ ���� ���ּ���');
        frm.gubuncode.focus;
        return;
    }
    <% end if %>
    
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
    
}
</script>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style>
.maintitle {color:red}
</style>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<div>
            <table width="660" cellpadding="2" cellspacing="1" class="tbType1 listTb">
            <form name="frmposcode" method="post" action="division_process.asp" >
            <input type="hidden" name="mode" value="<%=mode%>">
            <input type="hidden" name="gidx" value="<%=oDivision.FOneItem.Fgidx%>">
            <th colspan="2" style="padding:10px;">�̺�Ʈ ���� ���� - �׷����</th>
            <% if oDivision.FOneItem.Fgidx="" or oDivision.FOneItem.Fgidx=0 then %>
            <tr>
                <th width="200px;">�̺�Ʈ ���� ����</th>
                <td style="text-align:left;">
                    <select name="gubuncode">
                        <option value="1">���Ľ����̼�</option>
                        <option value="2">���� �̺�Ʈ</option>
                    </select>
                </td>
            </tr>
            <% end if %>
            <%'// ��ȹ�� selectBox %>
            <tr id="mastercode">
                <th>�ֻ��� ���� ���</th>
                <td style="text-align:left;"><%=DrawSelectAllView("mastercode",oDivision.FOneItem.Fmastercode,"")%> (��ü ���� �� �����ϸ� �ֻ��� ��� ����)</td>
            </tr>
            <tr>
                <th id="titlename">�����ڵ��</th>
                <td style="text-align:left;"><input type="text" name="title" value="<%=oDivision.FOneItem.Ftitle%>"/></td>
            </tr>
            <tr id="detailcode">
                <th>�����ڵ�</th>
                <td style="text-align:left;"><input type="text" name="detailcode" value="<%=oDivision.FOneItem.Fdetailcode%>" size="5"/>
                
                </td>    
            </tr>
            <tr>
                <th>��뿩��</th>
                <td style="text-align:left;">
                    <input type="radio" name="isusing" value="1" id="usey" <%=chkiif(oDivision.FOneItem.Fisusing = "" or oDivision.FOneItem.Fisusing = "1","checked","")%>><label for="usey">�����</label>
                    <input type="radio" name="isusing" value="0" id="usen" <%=chkiif(oDivision.FOneItem.Fisusing = "0","checked","")%>><label for="usen">������</div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center"><input type="button" value=" �� �� " onClick="SavePosCode(frmposcode);"></td>
            </tr>
            </form>
            </table>
            <%
            set oDivision = Nothing
            %>
            <br>
        </div>
        <div class="tPad15">
            <table width="660" cellpadding="2" cellspacing="1" class="tbType1 listTb">
            <tr>
                <td colspan="4" style="text-align:right"><a href="?gidx="><img src="/images/icon_new_registration.gif" border="0"></a></td>
            </tr>
            <tr>
                <th width="100">idx</th>
                <th>���и�</th>
                <th>�����ڵ��</th>
                <th>��뿩��</th>
            </tr>
            <% for i=0 to oDivisionList.FResultCount-1 %>
            <tr>
                <td><%= oDivisionList.FItemList(i).Fgidx %></td>
                <% if oDivisionList.FItemList(i).Fdetailcode < 0 then  %>
                <td style="text-align:left"><a href="?gidx=<%= oDivisionList.FItemList(i).Fgidx %>&page=<%= page %>"><span class="maintitle"><%=oDivisionList.FItemList(i).getGubunCodeName%></span></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oDivisionList.FItemList(i).Fgidx %>&page=<%= page %>"><span class="maintitle"><%=oDivisionList.FItemList(i).Ftitle%></span></a></td>
                <% else %>
                <td style="text-align:left"><a href="?gidx=<%= oDivisionList.FItemList(i).Fgidx %>&page=<%= page %>">&nbsp;��<%=getMasterCodeName(oDivisionList.FItemList(i).Fmastercode)%></a></td>
                <td style="text-align:left"><a href="?gidx=<%= oDivisionList.FItemList(i).Fgidx %>&page=<%= page %>">&nbsp;��<%=oDivisionList.FItemList(i).Ftitle%></a></td>
                <% end if %>
                <td><%= chkiif(oDivisionList.FItemList(i).Fisusing,"���","������") %></td>    
            </tr>
            <% next %>
            <tr>
                <td colspan="4" align="center">
                <% if oDivisionList.HasPreScroll then %>
                    <a href="?page=<%= oDivisionList.StartScrollPage-1 %>">[pre]</a>
                <% else %>
                    [pre]
                <% end if %>

                <% for i=0 + oDivisionList.StartScrollPage to oDivisionList.FScrollCount + oDivisionList.StartScrollPage - 1 %>
                    <% if i>oDivisionList.FTotalpage then Exit for %>
                    <% if CStr(page)=CStr(i) then %>
                    <font color="red">[<%= i %>]</font>
                    <% else %>
                    <a href="?page=<%= i %>">[<%= i %>]</a>
                    <% end if %>
                <% next %>

                <% if oDivisionList.HasNextScroll then %>
                    <a href="?page=<%= i %>">[next]</a>
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
    set oDivisionList = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->