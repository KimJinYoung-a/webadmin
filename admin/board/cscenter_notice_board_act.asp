<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.04.18 �ѿ�� 2008���� �̵�/�߰�
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/board/boardnoticecls.asp" -->
<%
dim boardnotice
dim boarditem
dim id, mode, title, contents, yuhyostart, yuhyoend, noticetype, fixyn, malltype
dim SearchKey, SearchString, param, page, listtype, menupos, retURL,oldyn, importantnotice

id = requestCheckVar(request("id"),12)
noticetype = requestCheckVar(request("noticetype"),8)
malltype = requestCheckVar(request("malltype"),8)
mode = requestCheckVar(request("mode"),128)

'����/������ ����ǥ�� ����Ѵ�.
title = request("title")
contents = request("contents")

yuhyostart = requestCheckVar(request("yuhyostart"),10)
yuhyoend = requestCheckVar(request("yuhyoend"),10)
fixyn = requestCheckVar(request("fixyn"),8)

page = requestCheckVar(request("page"),8)
SearchKey = requestCheckVar(request("SearchKey"),128)
SearchString = requestCheckVar(request("SearchString"),128)
listtype = requestCheckVar(request("listtype"),128)
menupos = requestCheckVar(request("menupos"),128)
oldyn = requestCheckVar(request("oldyn"),128)
importantnotice = requestCheckVar(request("importantnotice"),8)



param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&oldyn="& oldyn&"&noticetype=" & noticetype & "&menupos=" & menupos

if (mode = "write") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.Fnoticetype = noticetype
        boarditem.Ftitle = html2db(title)
        boarditem.Fcontents = html2db(contents)
        boarditem.Fyuhyostart = yuhyostart
        boarditem.Fyuhyoend = yuhyoend
        boarditem.Ffixyn = fixyn
        boarditem.Fmalltype = malltype
        boarditem.FImportantNotice = importantnotice
        boardnotice.write(boarditem)

        retURL = manageUrl & "/admin/board/" & "cscenter_notice_board_list.asp"
        'response.write "<script>location.replace('cscenter_notice_board_list.asp')</script>"
end if

if (mode = "modify") then
        set boardnotice = New CBoardNotice
        set boarditem = new CBoardNoticeItem

        boarditem.Fid = id
        boarditem.Fnoticetype = noticetype
        boarditem.Ftitle = html2db(title)
        boarditem.Fcontents = html2db(contents)
        boarditem.Fyuhyostart = yuhyostart
        boarditem.Fyuhyoend = yuhyoend
        boarditem.Ffixyn = fixyn
        boarditem.Fmalltype = malltype
        boarditem.FImportantNotice = importantnotice        
        boardnotice.modify(boarditem)

        retURL = server.URLencode(manageUrl & "/admin/board/" & "cscenter_notice_board_list.asp?page=" & page & param)
        'response.write "<script>location.replace('cscenter_notice_board_list.asp?page=" & page & param & "')</script>"
end if

if (mode = "delete") then
        set boardnotice = New CBoardNotice

        boardnotice.delete(id)

        retURL = server.URLencode(manageUrl & "/admin/board/" & "cscenter_notice_board_list.asp?page=" & page & param)
        'response.write "<script>location.replace('cscenter_notice_board_list.asp?page=" & page & param & "')</script>"
end if

'// ���� ���������� ��� �ε������� �̺�Ʈ���� '2009-04-18 �ѿ��
''if noticetype = 06 then
	'response.Redirect wwwUrl & "/chtml/make_index_culturestation.asp?retURL=" & retURL
''end if

'// �ε��� ������ HTML���� �������� �̵� (2008-01-10;������ ����)
response.Redirect wwwUrl & "/chtml/make_index_notice.asp?retURL=" & retURL
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->