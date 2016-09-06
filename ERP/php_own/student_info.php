<html>
<title>Sacs Student Corner</title>
<head><pre><center><font color="red" size="25">STUDENT INFORMATION RECORD</font></center></pre></head>
<body oncontextmenu="return false">
<div align="center">
	<form action="" method="POST" enctype="multipart/form-data">
 		REGISTER NUMBER&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="reg" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		NAME&nbsp;&nbsp;:&nbsp;&nbsp;<input type="text" name="name" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		DATE OF BIRTH&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="dob" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	
 		GENDER&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="gender" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		ADDRESS &nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="add" size="30"><br><br>
 		DEPARTMENT&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	 <select name="dept" >
	 <option value="COMPUTER SCIENCE">COMPUTER SCIENCE</option>
	 <option value="AUTOMOBILE">AUTOMOBILE</option>
	 <option value="CIVIL">CIVIL</option>
	 <option value="ELECTRICAL">ELECTRICAL</option>
	 <option value="ELECTRONICS">ELECTRONICS</option>
	 <option value="MECHANICAL">MECHANICAL</option>
	</select>
	 <br/><br>
 		CONTACT NUMBER&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="num" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		E-MAIL &nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="email" size="30"><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		BATCH &nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="text" name="batch" size="30"><br><br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		PHOTO &nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;<input type="file" name="image"/><br><br><br>
 			<input type="submit" name="submit" value="SAVE"/>
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
$register=$_POST['reg'];
$nam=$_POST['name'];
$dob=$_POST['dob'];
$gen=$_POST['gender'];
$addr=$_POST['add'];
$dept=$_POST['dept'];
$contact=$_POST['num'];
$email=$_POST['email'];
$batch=$_POST['batch'];
$image=addslashes($_FILES['image']['tmp_name']);
$name=addslashes($_FILES['image']['name']);
$image=file_get_contents($image);
$image=base64_encode($image);
if($dept=="COMPUTER SCIENCE")
{
	$sql="insert into cse_student_info values('$register','$nam','$dob','$gen','$addr','$dept','$contact','$email','$batch','$image')";
	if($conn->query($sql)==true)
	{
		echo "inserterd";
		$conn->close();
	}
	}
	else
	{
		echo "error in inserting";
		$conn->close();
	}
}
?>