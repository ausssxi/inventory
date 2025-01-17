<?php
ini_set('display_errors', "On");
$mysqli = new mysqli('localhost', 'root', 'password', 'merci_flower');
$mysqli->set_charset("utf8");
if (isset($_POST["course_name"])) {
    session_start();
    $_SESSION["course_name"] = $_POST["course_name"];
} elseif (isset($_POST["store_id"])) {
    $json = json_encode($_POST['quantity']);
    $result = $mysqli->prepare("UPDATE route SET $_POST[category] = ? WHERE store_id = ?");
    $result->bind_param('si', $json, $_POST['store_id']);
    $result->execute();
}
?>
