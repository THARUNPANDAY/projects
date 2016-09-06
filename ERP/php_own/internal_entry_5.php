<html>
<title>Sacs Student Corner</title>
<head><pre><center><font color="red" size="25">SACS STUDENT CORNER</font></center></pre></head>
<body oncontextmenu="return false">
	<div align="center">
		<form action="" method="POST">
	<font color="blue" size="5">
		REGISTER NUMBER &nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="reg" size="20"><br><br><br>
		<table border="0" cellspacing="0" >
			
			<tr><td>Subject code 1&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="sc1" size="10"></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Mark 1&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="m1" size="10"></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr>
			<tr><td>Subject code 2&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="sc2" size="10"></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Mark 2&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="m2" size="10"></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr>
			<tr><td>Subject code 3&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="sc3" size="10"></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Mark 3&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="m3" size="10"></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr>
			<tr><td>Subject code 4&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="sc4" size="10"></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Mark 4&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="m4" size="10"></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr>
			<tr><td>Subject code 5&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="sc5" size="10"></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>Mark 5&nbsp;&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="m5" size="10"></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr>
	    </table>
	    <br><br>
	    <input type="submit" name="submit" value="UPLOAD">
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
$conn = new mysqli($servername, $username, $password, $db);
// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
$reg=$_POST['reg'];
$sc1=$_POST['sc1'];
$sc2=$_POST['sc2'];
$sc3=$_POST['sc3'];
$sc4=$_POST['sc4'];
$sc5=$_POST['sc5'];
$m1=$_POST['m1'];
$m2=$_POST['m2'];
$m3=$_POST['m3'];
$m4=$_POST['m4'];
$m5=$_POST['m5'];
$sql="insert into cse_internal_marks values('$sc1','$sc2','$sc3','$sc4','$sc5','$m1','$m2','$m3','$m4','$m5','$reg','13')";
	if($conn->query($sql)==true)
	{
		echo "inserterd";
		$conn->close();
	}
	else
	{
		echo "error in inserting";
		$conn->close();
	}
}
?>