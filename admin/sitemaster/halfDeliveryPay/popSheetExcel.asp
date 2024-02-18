<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/halfDeliveryPay/halfdeliverypaycls.asp"-->
<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
        <meta name=ProgId content=Excel.Sheet>
        <meta name=Generator content="Microsoft Excel 9">
    </head>
    <body>
<%
    Dim loginUserId, i, currpage, pagesize, keyword, research, itemid, startdate, enddate, isusing, brandid, itemname, regusertype, regusertext
    Dim oHalfDeliveryPayList

    loginUserId = session("ssBctId") '// �α����� ����� ���̵�
    itemname = requestcheckvar(request("itemname"), 20) '// ��ǰ�� �˻���
    research = requestcheckvar(request("research"), 20) '// ��˻�����
    itemid = requestcheckvar(request("itemid"), 2048) '// ��ǰ�ڵ� �˻���
    startdate = requestcheckvar(request("startdate"), 20) '// ������ �˻���
    enddate = requestcheckvar(request("enddate"), 20) '// ������ �˻���
    isusing = requestcheckvar(request("isusing"), 20) '// ��뿩�� �˻���
    brandid = requestcheckvar(request("brandid"), 250) '// �귣�� ���̵� �˻���
    regusertype = requestcheckvar(request("regusertype"), 250) '// �ۼ��� �˻��ɼ�(id-���̵�, name-�̸�)
    regusertext = requestcheckvar(request("regusertext"), 250) '// �ۼ��� �˻� ��

    set oHalfDeliveryPayList = new CgetHalfDeliveryPay
        oHalfDeliveryPayList.FRectcurrpage = 1
        oHalfDeliveryPayList.FRectpagesize = 2000
        If Trim(research)="on" Then
            oHalfDeliveryPayList.FRectItemIds = itemid
            oHalfDeliveryPayList.FRectItemName = itemname
            oHalfDeliveryPayList.FRectStartdate = startdate
            oHalfDeliveryPayList.FRectEnddate = enddate
            oHalfDeliveryPayList.FRectIsUsing = isusing
            oHalfDeliveryPayList.FRectBrandId = brandid
            oHalfDeliveryPayList.FRectRegUserType = regusertype
            oHalfDeliveryPayList.FRectRegUserText = regusertext
        End If
        oHalfDeliveryPayList.GetHalfDeliveryPayList()

    If oHalfDeliveryPayList.FResultcount < 1 Then
        response.write "<script>alert('�����Ͱ� �����ϴ�.');window.close();</script>"
        response.end
    End If

	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=��ۺ�δ㼳������Ʈ" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
	Response.CacheControl = "public"

    'response.ContentType = "application/vnd.ms-excel"
    'Response.AddHeader "Content-Disposition", "attachment; filename=��ۺ�δ㼳������Ʈ_"& left(now(), 10) & ".xls"
    'response.write "<meta http-equiv=Content-Type content='text/html; charset=euc-kr'>"

    Function AddSpace(byval str)
        if ((str = "") or (IsNull(str))) then
            AddSpace = "&nbsp;"
        else
            AddSpace = str
        end if
    End Function

    function ConvertCurrencyUnit(str)
        if (str = "USD") then
            ConvertCurrencyUnit = "$"
        else
            ConvertCurrencyUnit = "��"
        end if
    End Function
%>
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
            <tr valign="top">
                <td class="td_br" width="80" align="center">��ȣ(idx)</td>
                <td class="td_br" width="100" align="center">��ǰ�ڵ�</td>
                <td class="td_br" width="100" align="center">�귣����̵�</td>
                <td class="td_br" align="center">��ǰ��</td>
                <td class="td_br" width="90" align="center">������</td>
                <td class="td_br" width="90" align="center">������</td>
                <td class="td_br" width="120" align="center">��۱���</td>
                <td class="td_br" width="170" align="center">�����ش��ǰ��۱���</td>                
                <td class="td_br" width="170" align="center">�����۱��رݾ�</td>
                <td class="td_br" width="100" align="center">��ۺ�</td>
                <td class="td_br" width="120" align="center">��ۺ�δ�ݾ�</td>
                <td class="td_br" width="80" align="center">��뿩��</td>
                <td class="td_br" width="240" align="center">�����</td>
                <td class="td_br" width="240" align="center">����������</td>
                <td class="td_br" width="120" align="center">�ۼ���</td>
                <td class="td_br" width="120" align="center">����������</td>                
            </tr>
            <% If oHalfDeliveryPayList.FResultcount > 0 Then %>
                <% For i=0 To oHalfDeliveryPayList.Fresultcount-1 %> 
                    <tr align="center" bgcolor="#FFFFFF">
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fidx%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemId%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fbrandid%></td>
                        <td class="td_br" align="left"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fitemname%></td>
                        <td class="td_br"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fstartdate,10)%></td>
                        <td class="td_br"><%=Left(oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fenddate,10)%></td>
                        <td class="td_br"><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliveryType)%></td>
                        <td class="td_br"><%=getBeadalDivname(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FItemDeliveryType)%></td>                        
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultFreeBeasongLimit,0)%>��</td>
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FdefaultDeliverPay,0)%>��</td>
                        <td class="td_br"><%=Formatnumber(oHalfDeliveryPayList.FHalfDeliveryPayList(i).FHalfDeliveryPay,0)%>��</td>
                        <td class="td_br">
                            <%
                                If oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fisusing = "Y" Then
                                    Response.write "���"
                                Else
                                    Response.write "������"
                                End If
                            %>                
                        </td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fregdate%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastupdate%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Fadminid%></td>
                        <td class="td_br"><%=oHalfDeliveryPayList.FHalfDeliveryPayList(i).Flastadminid%></td>                        
                    </tr>
                <% next %>
            <% End If %>
        </table>
    </body>
</html>
<%
    set oHalfDeliveryPayList = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
