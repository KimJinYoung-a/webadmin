<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/event/imageLinkCls.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<%
Dim mode : mode = requestCheckvar(request("mode"),4)
Dim masterIdx : masterIdx = requestCheckvar(request("masterIdx"),16)
dim x1 : x1 = requestCheckVar(Request("x1"),4)
dim y1 : y1 = requestCheckVar(Request("y1"),4)
dim x2 : x2 = requestCheckVar(Request("x2"),4)
dim y2 : y2 = requestCheckVar(Request("y2"),4)
dim didx : didx = requestCheckVar(Request("didx"),10)

If masterIdx = "" Then
    response.write "<script>alert('정상적인 경로로 접근해 주세요');history.back();</script>"
    response.end
End If
dim LinkURL
If didx <> "" Then
    dim oLinkDetailContents
    set oLinkDetailContents = new CimageLink
    oLinkDetailContents.FRectIdx = didx
    oLinkDetailContents.GetOneDetailContents()
    LinkURL=oLinkDetailContents.FOneItem.FLinkURL
    Set oLinkDetailContents = Nothing
end if
%>
<script>
    function SaveMapContents(mode){

        if(mode=="D"){
            if(confirm("설정된 맵 정보를 삭제 하시겠습니까?")){
                $.ajax({
                    type: "POST",
                    url: "/admin/eventmanage/event/v5/lib/ajaxEventImageLinkSet.asp",
                    data: "mode="+mode+"&masterIdx=<%=masterIdx%>&didx=<%=didx%>&x1=<%=x1%>&y1=<%=y1%>&x2=<%=x2%>&y2=<%=y2%>&linkurl="+$("#LinkURL").val(),
                    cache: false,
                    success: function(message) {
                        if(message=="0") {
                            alert("삭제 되었습니다.");
                            opener.location.reload();
                            self.close();
                        }
                        else if(message=="1"){
                            alert("유효하지 않은 데이터 입니다. 다시 시도해 주세요.");
                        }
                        else if(message=="2"){
                            alert("데이터 처리에 문제가 발생하였습니다.");
                        }
                        else if(message=="3"){
                            alert("유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.");
                        }
                    },
                    error: function(err) {
                        alert(err.responseText);
                    }
                });
            }
        }
        else{
            if($("#LinkURL").val()==""){
                alert("링크 정보를 먼저 입력해주세요.");
            }
            else{
                $.ajax({
                    type: "POST",
                    url: "/admin/eventmanage/event/v5/lib/ajaxEventImageLinkSet.asp",
                    data: "mode="+mode+"&masterIdx=<%=masterIdx%>&didx=<%=didx%>&x1=<%=x1%>&y1=<%=y1%>&x2=<%=x2%>&y2=<%=y2%>&linkurl="+$("#LinkURL").val(),
                    cache: false,
                    success: function(message) {
                        if(message=="0") {
                            alert("적용 되었습니다.");
                            opener.location.reload();
                            self.close();
                        }
                        else if(message=="1"){
                            alert("유효하지 않은 데이터 입니다. 다시 시도해 주세요.");
                        }
                        else if(message=="2"){
                            alert("데이터 처리에 문제가 발생하였습니다.");
                        }
                        else if(message=="3"){
                            alert("유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.");
                        }
                    },
                    error: function(err) {
                        alert(err.responseText);
                    }
                });
            }
        }
    }
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>이미지 맵 링크</h1>
	</div>
	<div class="popContV19">
		<table class="tableV19A">
			<colgroup>
				<col style="width:auto;">
			</colgroup>
			<tbody>
                <tr>
                    <td>
                        <input type="text" class="formControl" placeholder="링크URL" value="<%=LinkURL%>" name="LinkURL" id="LinkURL">
                    </td>
                </tr>
			</tbody>
        </table>
    </div>
	<div class="popBtnWrapV19">
        <% if mode="edit" then %>
		<button class="btn4 btnWhite1" onClick="SaveMapContents('D');">삭제</button>
        <button class="btn4 btnBlue1" onClick="SaveMapContents('E');return false;">저장</button>
        <% else %>
        <button class="btn4 btnBlue1" onClick="SaveMapContents('W');return false;">저장</button>
        <% end if %>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->