<html>
<title>Sacs Student Corner</title>
<head><pre><center><font color="red" size="25">SACS STUDENT CORNER</font></center></pre></head>
<body oncontextmenu="return false">
	<div align="center">
		<form action="" method="POST">
	<font color="blue" size="5">
			<br>Total Days&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="tot_days" size="5"><br><br>
			<table><tr><td><font color="blue" size="5">Register Number Of Absentees&nbsp;&nbsp;:</font></td><td><textarea rows="4" cols="50" name="regs"></textarea></td></tr></table><br>
			<input type="submit" name="submit" value="UPLOAD"/>
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
$regs=$_POST['regs'];
$tot_days=$_POST['tot_days'];
$conn = new mysqli($servername, $username, $password, $db);
// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
$st=0;

for($i=0;$i<=(strlen($regs)/12)-1;$i++)
{
	$result[$i]=substr($regs,$st,12);
	$st=$st+13;
	//echo $result[$i];
}
for($i=0;$i<count($result);$i++)
{
	$sql="SELECT * FROM cse_attendance WHERE register_number='$result[$i]'";
	$res = $conn->query($sql);
if ($res->num_rows > 0) {
		$date=strtolower(date("F"));
		$s_date=strtolower(date("M")."_tot");
		$sql1="SELECT $date FROM cse_attendance WHERE register_number='$result[$i]'";
		$a_result=$conn->query($sql1);
		$ab=$a_result->fetch_assoc();
		$tot_abs=$ab[$date].'a';
		$sql="update cse_attendance set $date ='$tot_abs',$s_date='$tot_days' where register_number='$result[$i]'";
		if($conn->query($sql)==true)
	  	{
    	echo "<center>"."UPDATED SUCCESSFULLY"."<center>";
    	} 
		else
	 	{
    	echo "<center>"."PROBLEM WHEN UPDATING"."</center>";
		}
		}
		else
		{
		echo "<center>"."WRONG REGISTER NUMBER"."</center>";
		}
}
}
?>