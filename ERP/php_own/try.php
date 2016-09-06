<?php
$servername = "localhost";
$username = "root";
$password = "";
$db="rrrr";

// Create connection
$conn = new mysqli($servername, $username, $password, $db);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
echo "Connected successfully";
$sql = "SELECT * FROM d";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
    // output data of each row
    while($row = $result->fetch_assoc()) {
        echo "id: " . $row["ll"]. " - Name: " . $row["ff"]. "<br>";
    }
} else {
    echo "0 results";
}
$conn->close();
?>
