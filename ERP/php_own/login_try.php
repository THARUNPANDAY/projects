<!DOCTYPE html>
<!--
Template Name: Academic Education V2
Author: <a href="http://www.os-templates.com/">OS Templates</a>
Author URI: http://www.os-templates.com/
Licence: Free to use under our free template licence terms
Licence URI: http://www.os-templates.com/template-terms
-->
<html>
<head>
<title>Sacs Student System</title>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<link href="layout/styles/layout.css" rel="stylesheet" type="text/css" media="all">
</head>
<body id="top">
<!-- ################################################################################################ --> 
<!-- ################################################################################################ --> 
<!-- ################################################################################################ -->
<div class="wrapper row0">
  <div id="topbar" class="clear"> 
    <!-- ################################################################################################ -->
  
    <!-- ################################################################################################ --> 
  </div>
</div>
<!-- ################################################################################################ --> 
<!-- ################################################################################################ --> 
<!-- ################################################################################################ -->
<div class="wrapper row1">
  <header id="header" class="clear"> 
    <!-- ################################################################################################ -->
    <div id="logo" class="fl_left">
      <h1><a href="index.html">Sacs Student System</a></h1>
      <p>a full student package</p>
    </div>
    
    <!-- ################################################################################################ --> 
  </header>
</div>
<!-- ################################################################################################ --> 
<!-- ################################################################################################ --> 
<!-- ################################################################################################ -->
<div class="wrapper row2">
  <div class="rounded">
    <nav id="mainav" class="clear"> 
      <!-- ################################################################################################ -->
      <ul class="clear">
        <li class="active"><a href="!">Admin Login</a></li>
        <li><a href="#">Staff Login</a></li>
      </ul>
      <!-- ################################################################################################ --> 
    </nav>
  </div>
</div>
<!-- ################################################################################################ --> 
<!-- ################################################################################################ --> 
<!-- ################################################################################################ -->
<div class="wrapper">
 
    <div  class="rounded clear"> 
      <!-- ################################################################################################ -->
          <form action="" method="post"><center>
  <font color="#000000" size="4">
  NAME : <input type="text" name="reg" size="25"/><br/><br/>
   PASSWORD:<input type="password" name="dob" size="25"/><br/><br/>
   DEPARTMENT:
   <select name="dept" >
   <option value="COMPUTER SCIENCE">COMPUTER SCIENCE</option>
   <option value="AUTOMOBILE">AUTOMOBILE</option>
   <option value="CIVIL">CIVIL</option>
   <option value="ELECTRICAL">ELECTRICAL</option>
   <option value="ELECTRONICS">ELECTRONICS</option>
   <option value="MECHANICAL">MECHANICAL</option>
  </select>
   <br/><br><br><br>

  <input type="submit"  name="submit"value="Log In"/>
  </font></center>
  </form>
      <!-- ################################################################################################ --> 
    </div>
  </div>
</div>
<!-- ################################################################################################ --> 
<!-- ################################################################################################ --> 
<!-- ################################################################################################ -->
<!-- JAVASCRIPTS --> 
<script src="layout/scripts/jquery.min.js"></script> 
<script src="layout/scripts/jquery.fitvids.min.js"></script> 
<script src="layout/scripts/jquery.mobilemenu.js"></script> 
<script src="layout/scripts/tabslet/jquery.tabslet.min.js"></script>
</body>
</html>
<?php
  if(isset($_POST['submit'])) {
// Create connection
$servername = "localhost";
$username = "root";
$password = "";
$db="ERP";
$date=date("Y-m-d");
$time=date("h:i:sA");
$conn = new mysqli($servername, $username, $password, $db);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
$user=$_POST['reg'];
$pass=$_POST['dob'];
$depart=$_POST['dept'];
$sql = "SELECT * FROM login where username='$user' and password='$pass' and department='$depart'";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
  $sql="insert into log values('$user','$pass','$depart','$date','$time')";
  if($conn->query($sql)==true)
  {
    header_remove();
    header("Location:/php_own/try.php");
    }
} else {
    echo "<center>"."PLEASE CHECK THE DETAILS ENTERED"."</center>";
}
$conn->close();

}
?> 