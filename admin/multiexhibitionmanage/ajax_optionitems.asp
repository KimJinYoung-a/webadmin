<%@ language="VBScript" %>
<% option Explicit %>
<% response.charset = "euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
    dim itemid , detailcode , addquery , query
    dim tophtml , addhtml

    itemid = requestCheckVar(request("itemid"),10)
    detailcode = requestCheckvar(request("detailcode"),100)

    if itemid = "" then 
        response.write "<script>alert('��ǰ �ڵ尡 �����ϴ�.');</script>"
        response.end
    end if 

    if detailcode <> "" then 
        addquery = "and detailcode in ("&detailcode&")"
    end if 

    query = " SELECT d.idx , d.gubuncode , d.itemid , d.optioncode , g.typename , i.optiontypename , i.optionname , d.isusing "&vbcrlf
    query = query & " FROM db_item.dbo.tbl_exhibition_item_detail as d WITH(NOLOCK) "&vbcrlf
    query = query & " INNER JOIN db_item.dbo.tbl_item_option as i WITH(NOLOCK) "&vbcrlf
    query = query & " on d.itemid = i.itemid and d.optioncode = i.itemoption "&vbcrlf
    query = query & " CROSS APPLY ( "&vbcrlf
    query = query & "   SELECT typename FROM db_item.dbo.tbl_exhibitionevent_groupcode WITH(NOLOCK) "&vbcrlf
    query = query & "   WHERE detailcode = d.detailcode "&vbcrlf
    query = query & " ) as g "&vbcrlf
    query = query & " where d.itemid = "& itemid & addquery
    rsget.Open query,dbget,1
    if not rsget.EOF  then
        tophtml = "<tr bgcolor='#BAB2B0'><td>�����ڵ�</td><td>��ǰ�ڵ�</td><td>�ɼ��ڵ�</td><td>�˻����͸�</td><td>��ǰ�ɼ�Ÿ��</td><td>��ǰ�ɼǸ�</td><td>�����Ȳ</td><td>����</td></tr>"
        do until rsget.EOF
            addhtml = addhtml & "<tr "& chkiif(rsget("isusing")=0 ,"bgcolor='#EC3F1A'","bgcolor='#FFFFFF'") &">"
            addhtml = addhtml & "<td>"& unescape(rsget("gubuncode")) &"</td>"
            addhtml = addhtml & "<td>"& unescape(rsget("itemid")) &"</td>"
            addhtml = addhtml & "<td>"& unescape(rsget("optioncode")) &"</td>"
            addhtml = addhtml & "<td>"& unescape(rsget("typename")) &"</td>"
            addhtml = addhtml & "<td>"& unescape(rsget("optiontypename")) &"</td>"
            addhtml = addhtml & "<td>"& unescape(rsget("optionname")) &"</td>"
            addhtml = addhtml & "<td><a href=""javascript:PopItemStock('"& rsget("gubuncode") &"','"& rsget("itemid") &"','"& rsget("optioncode") &"')"" title=""�����Ȳ �˾�"">[����]</a></td>"
            addhtml = addhtml & "<td id=""idx"& rsget("idx") &"""><a href=""javascript:FnIsUsing('"& rsget("idx") &"','"& chkiif(rsget("isusing")=1,"0","1") &"');"">[���"& chkiif(rsget("isusing")=1,"����","��") &"]</a></td>"
            addhtml = addhtml & "</tr>"
            
            rsget.MoveNext
        loop
        response.write tophtml & addhtml
    end if
    rsget.close    
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->