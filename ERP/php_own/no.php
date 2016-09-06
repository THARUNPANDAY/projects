
<html>
<body>
<form action="" method="post">
<input type="text" name="try"/>
<input type="submit" name="submit"/>
</form>
</body>
</html>

<?php
if(isset($_POST['submit']))
{
echo $_POST['try']; 
}
?>
