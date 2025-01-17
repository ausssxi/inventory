<?php
ini_set('display_errors', "On");
require_once(__DIR__ . '/vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;

$mysqli = new mysqli('localhost', 'root', 'password', 'merci_flower');
$mysqli->set_charset("utf8");
$result = $mysqli->prepare("UPDATE route SET turn = null, course_name = null;");
$result->execute();

$reader = new XlsxReader();
$spreadsheet = $reader->load('コース表.xlsx');
$sheet = $spreadsheet->getSheetByName('コース');
$data = $sheet->rangeToArray('B5:R100');
$course_name = [101, 102, 103, 104, 105, 201, 202, 203, 301, 302, 303, 304, 401, 402, 403, 404, 405, 501, 502, 503, 504, 505, 601, 602, 603, 604, 605];
for($i = 0; $i < 96; $i++) {
    for($j = 0; $j < 17; $j++) {
        foreach ($course_name as $value) {
            if ($data[$i][$j] == $value) {
                $course_order = 2;
                for ($ii = ($i + 2); $ii < ($i + 14); $ii++) {
                    $store_name = $data[$ii][$j];
                    if ($store_name !== "") {
                        $result = $mysqli->query("SELECT EXISTS(SELECT 1 FROM route WHERE store_name ='" . $value . "')");
                        $row = $result->fetch_row();
                        if ($row[0] == 0) {
                              $order_1 = 1;
                              $result = $mysqli->prepare("INSERT INTO route (store_name, turn, course_name) VALUES (?, ?, ?)");
                              $result->bind_param('sii', $value, $order_1, $value);
                              $result->execute();
                        }
                        $result = $mysqli->query("SELECT EXISTS(SELECT 1 FROM route WHERE store_name ='" . $store_name . "')");
                        $row = $result->fetch_row();
                        if ($row[0] == 0) {
                              $result = $mysqli->prepare("INSERT INTO route (store_name, course_name) VALUES (?, ?)");
                              $result->bind_param('si', $store_name, $value);
                              $result->execute();
                        }
                        $result = $mysqli->prepare("UPDATE route SET turn = ?, course_name = ? WHERE store_name = ?");
                        $result->bind_param('iis', $course_order, $value, $store_name);
                        $result->execute();
                    }
                    $course_order++;
                }
            }
        }
    }
}
$int = 1;
foreach ($course_name as $value) {
    $result = $mysqli->prepare("UPDATE route SET turn = ?, course_name = ? WHERE store_name = ?");
    $result->bind_param('iis', $int, $value, $value);
    $result->execute();
}
$mysqli->close();
?>
