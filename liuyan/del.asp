<!--
---------------------文    件:del.asp
---------------------程序名字:留言板程序(附全注释)
---------------------笔者网名:Dmlt
---------------------笔者博客:Http://HI.Baidu.com/Int_yi
---------------------笔者附言:如果有BUG,或者漏写部分,请于与人联系;
---------------------联系方式:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<!--#include file="sqlin.asp" -->
<%
    if session("pass")="" then
	response.write "还没有登陆验证呢!"
	response.end
    end if
%>

<%
id=request("id")                                        ' 获取ID
conn.execute("delete from data where id in ("&id&")")   '删除查询到的ID
response.write "<script>history.go(-1)</script>"        '返回到上页
%>