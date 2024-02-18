<%
function httpBuildQueryArray(key, items, count, out)
    dim result
    for i = 0 to count - 1
        if isarray(items(i)) then
            httpBuildQueryArray (key & "[" & i & "]"), items(i), UBound(items(i)), result
        else
            result = result & Server.URLEncode(key & "[" & i & "]") & "=" & Server.URLEncode(items(i))
        end if
        if i + 1 <> count then
            result = result & "&"
        end if
    next
    out = result
end function
' httpBuildQuery(
'   keys array     The keys of hash.Keys
'   items array    The items of hash.Items
'   count integer  The count of hash.Count
'   out string     The output of the string
' )
' Takes a Scripting.Dictionary and build a query string
function httpBuildQuery(keys, items, count, out)
    dim result
    for i = 0 to count - 1
        if isarray(items(i)) then
            httpBuildQueryArray keys(i), items(i), UBound(items(i)), result
        else
            result = result & Server.URLEncode(keys(i)) & "=" & Server.URLEncode(items(i))
        end if
        if i + 1 <> count then
            result = result & "&"
        end if
    next
    out = result
end function
%>