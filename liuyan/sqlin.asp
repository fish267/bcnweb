

<%
id=trim(request("id"))                                '取得ID后面的号码,为后面判断关键词.
if id="" or instr(id,"'")>0  or instr(id," ")>0 then  '过滤掉SQL的敏感词，特别是空格和 单撇号
    response.write "内容提交有误!"
    response.end
end if 
%>