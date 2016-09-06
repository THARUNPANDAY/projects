<html>
<title>Sacs Student Corner</title>
<head><pre><center><font color="red" size="25">SACS STUDENT CORNER</font></center></pre></head>
<body oncontextmenu="return false">
	<div align="center">
		<form action="" method="post">
	<font color="blue" size="5">
	NAME &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="reg" size="25"/><br/><br/>
	 PASSWORD&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;<input type="password" name="dob" size="25"/><br/><br/>
	 DEPARTMENT&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
	</font>
	</form>
	</div>
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
	$sql="insert into log values('$user','$pass','$dapart','$date','$time')";
	if($conn->query($sql)==true)
	{
    header("Location:/php_own/student_info.php");
    }
} else {
    echo "<center>"."PLEASE CHECK THE DETAILS ENTERED"."</center>";
}
$conn->close();
}
?> 