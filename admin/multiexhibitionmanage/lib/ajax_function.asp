<%@ language="VBScript" %>
<% option Explicit %>
<% response.charset = "euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
    dim mastercode , filtercode , targetform , targetname
    dim query

    mastercode = requestCheckVar(request("mastercode"),10)
    filtercode = requestCheckVar(request("filtercode"),1)
    targetform = requestCheckVar(request("targetform"),10)
    targetname = requestCheckVar(request("targetname"),10)

    '// ��ũ��Ʈ ����
    function mkjsfuntion(frm,frmname)
        if frm <> "" and frmname <> "" then 
            mkjsfuntion = "document."& frm &"."& targetname &".value = this.value;"
        end if 
    end function

    '// ���� ó��
    if mastercode = "" then
        response.write "<script>alert('��ȹ�� �ڵ尡 �����ϴ�.');</script>"
        response.end
    end if 

    '// filter code 
    '// 1.button type
    '// 2.checkbox type
    '// 3.radio type
    select case filtercode
        case "1"
            query = " select distinct type from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and mastercode = '"& mastercode &"' and detailcode > 0 order by type asc "
            rsget.Open query,dbget,1
            if not rsget.EOF  then
                do until rsget.EOF
                    response.write "<input type=""button"" class=""button_s"" value="""& unescape(rsget("type")) &""" onclick='"& mkjsfuntion(targetform,targetname) &"'>&nbsp;"
                    rsget.MoveNext
                loop
            else
                response.write "<script>alert('���и��� �����ϴ�.');document."& targetform &"."& targetname &".focus();</script>"
            end if
            rsget.close
        case "2", "3"
            dim buttiontype : buttiontype = chkiif(filtercode = 2,"checkbox" ,"radio")
            dim temptype , i

            query = " select distinct type , typename , detailcode from db_item.dbo.tbl_exhibitionevent_groupcode where isusing = 1 and mastercode = '"& mastercode &"' and detailcode > 0 order by type asc "
            rsget.Open query,dbget,1
            i = 0
            if not rsget.EOF  then
                do until rsget.EOF
                    if cstr(temptype) <> unescape(rsget("type")) then
                        if temptype <> "" then response.write "<br/><br/>"
                        response.write "<span style='width:100px; margin:0 auto; padding:10px; text-align:center; color:red'>"& unescape(rsget("type")) &"</span>&nbsp;"
                    end if

                    response.write "[ <input type="""& buttiontype &""" id='"& rsget("detailcode") & i &"' name=""detailcode"" class=""button_s"" value="""& rsget("detailcode") &""" onclick='"& chkiif(filtercode=3,mkjsfuntion(targetform,targetname),"void(0);") &"'> <label for='"& rsget("detailcode") & i &"' onclick='"& chkiif(filtercode=3,mkjsfuntion(targetform,targetname),"void(0);") &"'>"&unescape(rsget("typename")) &"</label> ]&nbsp;"

                    temptype = rsget("type")
                    i = i + 1
                    rsget.MoveNext
                loop
                
            else
                response.write "<script>alert('���и��� �����ϴ�.');document."& targetform &"."& targetname &".focus();</script>"
            end if
            
            rsget.close
        case else
            response.write "<script>alert('�����ڵ� ���� �־� �ּ���');</scirpt>"
    end select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->