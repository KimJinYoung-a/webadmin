<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim idxlist
dim obarcode

idxlist = request("idxlist")

set obarcode = new CShopIpChul
obarcode.FRectIdxArr = idxlist
obarcode.GetBaCodeListByIdxList

dim i
%>
<OBJECT
	  id=iaxobject
	  classid="clsid:5D776FEA-8C6B-4C53-8EC3-3585FC040BDB"
	  codebase="http://webadmin.10x10.co.kr/common/cab/tenbarPrint.cab#version=1,0,0,29"
	  width=0
	  height=0
	  align=center
	  hspace=0
	  vspace=0
>
</OBJECT>

<script language='javascript'>
function AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno){
    if (itemid*1>=1000000){
        iaxobject.AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype*10, itemno);
    }else{
	    iaxobject.AddData(itemid, itemoption, itemname, itemoptionname, brand, itemprice, itemtype, itemno);
	}
}
</script>
<script language='javascript'>
iaxobject.ClearItem();
//iaxobject.setTitleVisible(true);
<% for i=0 to obarcode.FresultCount -1 %>
AddData("<%= obarcode.FItemList(i).Fitemid %>",
"<%= obarcode.FItemList(i).Fitemoption %>",
"<%= Replace(obarcode.FItemList(i).Fitemname,chr(34),"") %>",
"<%= obarcode.FItemList(i).Fitemoptionname %>",
"<%= obarcode.FItemList(i).Fbrand %>",
"<%= obarcode.FItemList(i).Fitemprice %>",
"<%= obarcode.FItemList(i).Fitemtype %>",
"<%= obarcode.FItemList(i).Fitemno %>"
);
<% next %>
iaxobject.ShowFrm();
</script>
<%
set obarcode = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->