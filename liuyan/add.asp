<!--
---------------------文    件:add.asp
---------------------程序名字:留言板程序(附全注释)
---------------------笔者网名:Dmlt
---------------------笔者博客:Http://HI.Baidu.com/Int_yi
---------------------笔者附言:如果有BUG,或者漏写部分,请与本人联系;
---------------------联系方式:dmlt@vip.qq.com
-->

<!--#include file="conn.asp" -->
<%
name=trim(request("name"))                                                 '获取name,并且去处后面的空格(trim 函数出去空格)
if name=""  then                                                           '如果name的值为空,则输出一下内容
response.write "<script>alert('没有昵称');location='index.asp';onclick=history.go(-1)</script>"
    response.end                                                           '并且结束整个过程
end if 
ly=trim(request("ly"))                                                    '获取ly,并且去处后面的空格(trim 函数出去空格)
if ly=""  then                                                            '如果ly的值为空,则输出一下内容
response.write "<script>alert('没有内容');location='index.asp';onclick=history.go(-1)</script>"
    response.end                                                          '并且结束整个过程
end if                                                                    '结束if 语句
%>

<%
set rs=server.createobject("adodb.recordset")                              '对rs赋予权限
sql="select * from data"                                                   '对表的操作
rs.open sql,conn,3,3                                                       '操作数据库
rs.addnew                                                                  '添加新数据
rs("ly")=request.form("ly")                                                '赋值给ly,给数据库
rs("name")=request.form("name")                                            '赋值给name,给数据库
rs.update                                                                  'rs为update
rs.close                                                                   '结束一切rs操作
set rs=nothing                                                             '清空rs
conn.close                                                                 '结束一切conn操作
set conn=nothing                                                           '清空conn
response.write "<script>history.go(-1)</script>"                           '返回上一页
%>