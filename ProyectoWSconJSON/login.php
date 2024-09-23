<?php
$usua = $_REQUEST['USUARIO'];
$pass= $_REQUEST['PASSWORD'];

$mysqli = new mysqli('bdprueba.db.10193154.a06.hostedresource.net', 'bdprueba', 'Abcd123!', 'bdprueba');
$myArray = array();
if($result = $mysqli->query("SELECT * from TABLA1 where USUARIO='$usua' and PASSWORD='$pass'")){
    while ($row = $result->fetch_array(MYSQLI_ASSOC)){
        $myArray[] = $row;
    }
    echo json_encode($myArray);
}
?>