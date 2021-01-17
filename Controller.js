//<!--#include file="adojavas.inc"-->
function Connect(user, password) {
  var connectionString = 'Data Source=DESKTOP-88339Q2\\SQLEXPRESSIT350; Initial Catalog= Northwind; User ID=' + user + ';Password=' + password + ';Provider=SQLOLEDB';

  //alert(connectionString);

  var connection = new ActiveXObject('ADODB.Connection');
  connection.Open(connectionString);

  alert("Connected");
};

function CountRecords(user, password) {
  var connectionString = 'Data Source=DESKTOP-88339Q2\\SQLEXPRESSIT350;Initial Catalog=northwind;User ID=' + user + ';Password=' + password + ';Provider=SQLOLEDB';
  var connection = new ActiveXObject("ADODB.Connection");
  connection.Open(connectionString);

  var rs = new ActiveXObject("ADODB.Recordset");
  rs.Open("Select count(*) as NumRecords From Customers", connection);
  rs.MoveFirst
while (!rs.eof) {
    alert(rs.fields(0));
    rs.movenext;
  }
  rs.close;
  connection.close;
};
function ShowNames(user, password) {
  var connectionString = 'Data Source=DESKTOP-88339Q2\\SQLEXPRESSIT350;Initial Catalog=northwind;User ID=' + user + ';Password=' + password + ';Provider=SQLOLEDB';

  var connection = new ActiveXObject("ADODB.Connection");
  connection.Open(connectionString);
  var rs = new ActiveXObject("ADODB.Recordset");
  rs.Open("Select CompanyName From Customers", connection);
  rs.MoveFirst
  var input = ' ';
  while (!rs.eof) {
    input += rs.fields(0) + "\n";
    rs.movenext;
  }
  alert (input);

  rs.close;
  connection.close;
};
