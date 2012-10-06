<%
Dim Conn,ConnStr 
dim Sql_Server,Sql_User,Sql_Pass,Sql_Data 
Sql_Server = "127.0.0.1" '数据库服务器地址 
Sql_User = "sa" '数据库登录名 
Sql_Pass = "521521" '数据库密码 
Sql_Data = "uu00_dj" '数据库名
connstr = "PROVIDER=SQLOLEDB;DATA SOURCE="&Sql_Server&";UID="&Sql_User&";PWD="&Sql_Pass&";DATABASE="&Sql_Data 
set conn = server.createobject("ADODB.connection") '创建数据库连接对象 
conn.open connstr '连接数据库
%>

