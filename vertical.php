<?php
ini_set('display_errors', "On");
session_start();
require_once(__DIR__ . '/vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup as PageSetup;
$mysqli = new mysqli('localhost', 'root', 'password', 'merci_flower');
$mysqli->set_charset("utf8");

function upload($category, $line) {
    global $mysqli;
    global $spreadsheet;
    $day = 1;
    for ($sheet_no = 0; $sheet_no < 2; $sheet_no++) {
        $Row = 12;
        $sheet = $spreadsheet->getSheet($sheet_no);
        $count = 1;
        for ($i = 12; $i < 77; $i++) {
            if($i % 4 != 0) {
                continue;
            }
            $sheet->setCellValueByColumnAndRow(1, $i, $count);
            $count++;
        }
        $result = $mysqli->query("SELECT store_name FROM route WHERE course_name = $_SESSION[course_name] AND store_name != $_SESSION[course_name] order by turn");
        while ($row = $result->fetch_assoc()) {
            $sheet->setCellValueByColumnAndRow(2, $Row, $row["store_name"]);
            $Row += 4;
        }
        $sheet->setCellValue('B1', $_SESSION['course_name']);
        $sheet->setCellValue('B4', "棚卸");
        $sheet->setCellValue('B5', "${day}日分");
        $sheet->setCellValue('I4', "クリザール");
        $sheet->setCellValue('B80', "コース合計");
        $sheet->setCellValue('I83', "=SUM(I11,I15,I19,I23,I27,I31,I35,I39,I43,I47,I51,I55,I59,I63,I67,I71,I75,I79)");
        $items = ["店舗名", "合計", "資材", "切花", "園芸", "榊", "内訳"];
        $Column = 2;
        foreach ($items as $item) {
            $sheet->setCellValueByColumnAndRow($Column, 6, $item);
            $Column++;
        }
        $prices = [380, 80, 100, 120, 128, 158, 178, 198, 200, 228, 258, 298, 358, 398, 458, 498, 550, 598, 658, 698, 758, 798, 858, 898, 958, 980, 1280, 1580, 1980, 2580, 2980];
        $Column = 9;
        foreach ($prices as $price) {
            $sheet->setCellValueByColumnAndRow($Column, 6, $price);
            $Column++;
        }
        for ($i = 9; $i < 40; $i++) {
            $coordinate = Coordinate::stringFromColumnIndex($i);
            $sheet->setCellValueByColumnAndRow($i, 7, "=ROUND(${coordinate}6*1.1,0)");
        }
        $sheet->setCellValue('B8', "前月末日　戻り分");
        for ($i = 8; $i < 77; $i++) {
            if($i % 4 != 0) {
                continue;
            }
            $sheet->setCellValueByColumnAndRow(3, $i, "=SUM(D${i}:G${i})");
        }
        for ($i = 8; $i < 77; $i++) {
            if($i % 4 != 0) {
                continue;
            }
            $i_1 = $i + 1;
            $i_2 = $i + 2;
            $i_3 = $i + 3;
            $sheet->setCellValueByColumnAndRow(4, $i, "=SUMPRODUCT(\$I$6:\$AM$6,\$I${i_3}:\$AM${i_3})");
            $sheet->setCellValueByColumnAndRow(5, $i, "=SUMPRODUCT(\$J$6:\$AM$6,\$J${i}:\$AM${i})");
            $sheet->setCellValueByColumnAndRow(6, $i, "=SUMPRODUCT(\$J$6:\$AM$6,\$J${i_1}:\$AM${i_1})");
            $sheet->setCellValueByColumnAndRow(7, $i, "=SUMPRODUCT(\$J$6:\$AM$6,\$J${i_2}:\$AM${i_2})");
        }
        $items = ["切花", "園芸", "榊", "資材"];
        $Row = 8;
        for ($i = 0; $i < 19; $i++) {
            foreach ($items as $item) {
                $Row = $Row;
                $sheet->setCellValueByColumnAndRow(8, $Row, $item);
                $Row++;
            }
        }
        for ($i = 3; $i < 8; $i++) {
            $coordinate = Coordinate::stringFromColumnIndex($i);
            $sheet->setCellValueByColumnAndRow($i, 80, "=SUM(${coordinate}12:${coordinate}77)");
        }
        for ($i = 10; $i < 40; $i++) {
            $coordinate = Coordinate::stringFromColumnIndex($i);
            $sheet->setCellValueByColumnAndRow($i, 80, "=SUM(${coordinate}8,${coordinate}12,${coordinate}16,${coordinate}20,${coordinate}24,${coordinate}28,${coordinate}32,${coordinate}36,${coordinate}40,${coordinate}44,${coordinate}48,${coordinate}52,${coordinate}56,${coordinate}60,${coordinate}64,${coordinate}68,${coordinate}72,${coordinate}76)");
            $sheet->setCellValueByColumnAndRow($i, 81, "=SUM(${coordinate}9,${coordinate}13,${coordinate}17,${coordinate}21,${coordinate}25,${coordinate}29,${coordinate}33,${coordinate}37,${coordinate}41,${coordinate}45,${coordinate}49,${coordinate}53,${coordinate}57,${coordinate}61,${coordinate}65,${coordinate}69,${coordinate}73,${coordinate}77)");
            $sheet->setCellValueByColumnAndRow($i, 82, "=SUM(${coordinate}11,${coordinate}14,${coordinate}18,${coordinate}22,${coordinate}26,${coordinate}30,${coordinate}34,${coordinate}38,${coordinate}42,${coordinate}46,${coordinate}50,${coordinate}54,${coordinate}58,${coordinate}62,${coordinate}66,${coordinate}70,${coordinate}74,${coordinate}78)");
            $sheet->setCellValueByColumnAndRow($i, 83, "=SUM(${coordinate}12,${coordinate}15,${coordinate}19,${coordinate}23,${coordinate}27,${coordinate}31,${coordinate}35,${coordinate}39,${coordinate}43,${coordinate}47,${coordinate}51,${coordinate}55,${coordinate}59,${coordinate}63,${coordinate}67,${coordinate}71,${coordinate}75,${coordinate}79)");
        }
        for ($i = 1; $i < 9; $i++) {
            $coordinate = Coordinate::stringFromColumnIndex($i);
            $sheet->mergeCells($coordinate . "6:" . $coordinate . "7");
            if ($i != 8) {
                for ($j = 8; $j < 84; $j++) {
                    if($j % 4 != 0) {
                        continue;
                    }
                    $jj = $j + 3;
                    $sheet->mergeCells($coordinate . $j. ":" . $coordinate . $jj);
                }
            }
        }


        $borders = $sheet->getStyle('A6:AM83')->getBorders();
        $borders->getTop()->setBorderStyle('double');
        $borders->getBottom()->setBorderStyle('double');
        $borders->getRight()->setBorderStyle('double');
        $borders = $sheet->getStyle('A6:A83')->getBorders();
        $borders->getRight()->setBorderStyle('medium');
        for ($i = 2; $i < 39; $i++) {
            $coordinate = Coordinate::stringFromColumnIndex($i);
            $borders = $sheet->getStyle($coordinate . "6:" . $coordinate . "83")->getBorders();
            $borders->getRight()->setBorderStyle('thin');
        }

        for ($i = 7; $i < 84; $i++) {
            if ($i == 12) {
                $borders = $sheet->getStyle('A12:AM12')->getBorders();
                $borders->getTop()->setBorderStyle('double');
            } elseif($i == 80) {
                $borders = $sheet->getStyle('A80:AM80')->getBorders();
                $borders->getTop()->setBorderStyle('double');
            } elseif($i % 4 == 0) {
                $borders = $sheet->getStyle("A" . $i . ":AM" . $i)->getBorders();
                $borders->getTop()->setBorderStyle('thin');
            } else {
                $borders = $sheet->getStyle("I" . $i . ":AM" . $i)->getBorders();
                $borders->getTop()->setBorderStyle('hair');
            }
        }
        $sheet->getColumnDimension('B')->setWidth(20);
        for ($i = 9; $i < 47; $i++) {
            for ($j = 6; $j < 8; $j++) {
                $coordinate = Coordinate::stringFromColumnIndex($i);
                $sheet -> getStyle($coordinate . $j) -> getAlignment() -> setTextRotation(-165);
            }
        }
        $sheet->getStyle("A4:AT83")->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $sheet->getStyle("A4:AT83")->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
        $sheet->getStyle('B4:B5')->getFont()->setBold(true);
        $sheet->getStyle('B5')->getFont()->getColor()->setARGB('FFFF0000');
        for ($j = 8; $j < 83; $j++) {
            if($j % 4 == 3) {
                continue;
            }
            $coordinate = Coordinate::stringFromColumnIndex(9);
            $sheet->getStyle($coordinate . $j)->getFill()->setFillType('solid')->getStartColor()->setARGB('FF808080');
        }
        $sheet->getColumnDimension('D')->setVisible(false);
        $sheet->getColumnDimension('E')->setVisible(false);
        $sheet->getColumnDimension('F')->setVisible(false);
        $sheet->getColumnDimension('G')->setVisible(false);
        $sheet->getColumnDimension('H')->setVisible(false);
        for ($i = 48; $i < 80; $i++) {
            $sheet->getRowDimension($i)->setVisible(false);
        }
        $sheet->getPageSetup()->setPrintArea('A1:AT83');
        $sheet->getPageSetup()->setFitToPage(true);
        $sheet -> getPageSetup() -> setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        $category_day = $category.$day;
        $j = $line;
        $result = $mysqli->query("SELECT ${category_day} FROM route WHERE course_name = $_SESSION[course_name] order by turn");
        while ($row = $result->fetch_assoc()) {
            $quantity = json_decode($row[$category_day], true);
            $data = $sheet->rangeToArray('J6:AM6');
            for($i = 0; $i < 30; $i++) {
                foreach ((array)$quantity as $price => $value) {
                    if ($data[0][$i] == $price) {
                        $column = $i;
                        $column += 10;
                        $sheet->setCellValueByColumnAndRow($column, $j, $value);
                    } else if ($price === "クリザ") {
                        $column = 9;
                        $sheet->setCellValueByColumnAndRow($column, $j, $value);
                    }
                }
            }
            $j += 4;
        }
        $day += 15;
    }
}

if (isset($_POST["upload"])) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('1日');
    $day16 = new Worksheet($spreadsheet, '16日');
    $spreadsheet->addSheet($day16, 1);
    upload("cutflowers", 8);
    upload("horticulture", 9);
    upload("cleyera", 10);
    upload("materials", 11);

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $_SESSION['course_name'] . '.xlsx"');
    header('Cache-Control: max-age=0');
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
    die;
}
?>
<html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no">
        <link rel="shortcut icon" href="favicon.ico">
        <title>棚卸し</title>
        <script type="text/javascript" src="/jquery-3.3.1.min.js"></script>
    </head>
    <style>
        .flex {
            display: flex;
        }

        .fixed {
            position: fixed;
        }

        .bottom-0 {
            bottom: 0;
        }

        .width-100pe {
            width: 100%;
        }

        .height-48 {
            height: 48px;
        }

        .padding-top-16 {
            padding-top: 16px;
        }

        .padding-bottom-16 {
            padding-bottom: 16px;
        }

        .font-size-14 {
            font-size: 14px;
        }

        .text-align-center {
            text-align: center;
        }

        .background-lightgray {
            background: #d3d3d3;
        }

        .background-lightblue {
            background: #add8e6;
        }

        .background-lightgreen {
            background: #90ee90;
        }

        .background-lightyellow {
            background: #ffffe0;
        }

        .background-lightsalmon {
            background: #ffa07a;
        }

        .background-lightpink {
            background: #ffb6c1;
        }

        .background-white {
            background:#fff;
        }

        @media screen and (max-width: 767px) {
            .none-767 {
                display: none;
            }

            .width-16pe {
                width: 16%;
            }

            .width-18pe {
                width: 18%;
            }

            .width-19pe {
                width: 19%;
            }

            .width-21pe {
                width: 21%;
            }
        }

        @media screen and (min-width: 768px) {
            .space-between {
                justify-content: space-between;
            }
        }

        .tab-wrap {
	          background: White;
	          box-shadow: 0 0 5px rgba(0,0,0,.1);
	          display: flex;
	          flex-wrap: wrap;
	          overflow: hidden;
	          padding: 0 0 16px;
        }

        .tab-label {
	          color: Gray;
	          cursor: pointer;
	          flex: 1;
	          font-weight: bold;
	          order: -1;
	          padding: 16px;
	          position: relative;
	          text-align: center;
	          transition: cubic-bezier(0.4, 0, 0.2, 1) .2s;
	          user-select: none;
	          white-space: nowrap;
          	-webkit-tap-highlight-color: transparent;
        }

        .tab-label:hover {
	          background: rgba(0, 191, 255,.1);
        }

        .tab-switch:checked + .tab-label {
	          color: DeepSkyBlue;
        }

        .tab-label::after {
	          background: DeepSkyBlue;
	          bottom: 0;
	          content: '';
	          display: block;
	          height: 3px;
	          left: 0;
	          opacity: 0;
	          pointer-events: none;
           	position: absolute;
	          transform: translateX(100%);
	          transition: cubic-bezier(0.4, 0, 0.2, 1) .2s 80ms;
	          width: 100%;
	          z-index: 1;
        }

        .tab-switch:checked ~ .tab-label::after {
	          transform: translateX(-100%);
        }

        .tab-switch:checked + .tab-label::after {
	          opacity: 1;
	          transform: translateX(0);
        }

        .tab-content {
	          height:0;
	          opacity:0;
	          padding: 0 8px;
	          pointer-events:none;
	          transform: translateX(-30%);
	          transition: transform .3s 80ms, opacity .3s 80ms;
	          width: 100%;
         }

         .tab-switch:checked ~ .tab-content {
          	transform: translateX(30%);
         }

         .tab-switch:checked + .tab-label + .tab-content {
	          height: auto;
	          opacity: 1;
	          order: 1;
	          pointer-events:auto;
	          transform: translateX(0);
         }

         .tab-wrap::after {
	          content: '';
	          height: 16px;
	          order: -1;
	          width: 100%;
         }

         .tab-switch {
	          display: none;
         }

         <?php
         if (isset($_SESSION["course_name"])) {
             $result = $mysqli->query("SELECT count(*) FROM route WHERE course_name = $_SESSION[course_name]");
             $row = $result->fetch_array();
             $width = $row[0] * 208;
             $width_06 = $width * 0.6;
         }
         ?>

         .tab-container, .tab-content {
	           max-width: calc(100vw - 16px*2);
             <?php
             if (isset($_SESSION["course_name"])) {
	               echo "width:${width_06}px";
             }
             ?>
         }

         .tab-wrap {
             <?php
             if (isset($_SESSION["course_name"])) {
	               echo "width:${width}px";
             }
             ?>
         }

         @media screen and (max-width: 767px) {
             .width-100-06 {
                 width: 100%;
              }
         }

         @media screen and (min-width: 768px) {
             .width-100-06 {
                 <?php
                 if (isset($_SESSION["course_name"])) {
                     echo "width:${width_06}px";
                 }
                 ?>
             }
         }

         * {
	           box-sizing: border-box;
         }

         .tab-container {
	           box-shadow: 0 0 5px rgba(0,0,0,.1);
             overflow: hidden;
	           overflow-x: auto;
	           position: relative;
         }
         .tab-wrap {
	           box-shadow: none;
	           overflow: visible;
         }
         .tab-content {
	           left:0;
             position: -webkit-sticky;
	           position: sticky;
         }

         .tab-wrap::before {
	           content: '';
	           height: 0;
	           order: 1;
	           width: 100%;
         }

         body {
	           background: WhiteSmoke;
	           font-family: sans-serif;
	           margin: 16px;
         }
    </style>
    <body>
        <?php
        if(!empty($_GET['logout'])) {
            unset($_SESSION['course_name']);
            unset($_POST['course_name']);
        }
        if (isset($_SESSION["course_name"])) {
            echo "<div class='tab-container'>
                <div class='tab-wrap'>";
                    $i = 1;
                    $result_course = $mysqli->query("SELECT store_id, store_name FROM route WHERE course_name = $_SESSION[course_name] order by turn");
                    while ($row = $result_course->fetch_assoc()) {
                        $store_id = $row["store_id"];
                        if(mb_strlen($row["store_name"]) > 10) {
                            $store_name = mb_substr($row["store_name"] , 0, 10);
                            $store_name = $store_name . "…";
                        } else {
                            $store_name = $row["store_name"];
                            if (is_numeric($store_name)) {
                                $store_name = "戻り";
                            }
                        }
                        echo "<input id='TAB02-0${i}' type='radio' name='TAB02' class='tab-switch'";
                        if ($i == 1) {
                            echo " checked='checked'";
                        }
                        echo " /><label class='tab-label' for='TAB02-0${i}'>${store_name}</label>
                        <div class='tab-content'>";
                        if (8 < date("d") && date("d") < 24) {
                            $day = 16;
                        } else {
                            $day = 1;
                        }
                            echo "<table class='width-100pe'>";
                                $prices = ["クリザ", 80, 100, 120, 128, 158, 178, 198, 200, 228, 258, 298, 358, 398, 458, 498, 550, 598, 658, 698, 758, 798, 858, 898, 958, 980, 1280, 1580, 1980, 2580, 2980];
                                foreach ($prices as $price) {
                                    if ($price == 298) {
                                        echo "<tr class='background-lightgray'>";
                                    } else if ($price == 358) {
                                        echo "<tr class='background-lightblue'>";
                                    } else if ($price == 398) {
                                        echo "<tr class='background-lightgreen'>";
                                    } else if ($price == 458) {
                                        echo "<tr class='background-lightyellow'>";
                                    } else if ($price == 498) {
                                        echo "<tr class='background-lightsalmon'>";
                                    } else if ($price == 598) {
                                        echo "<tr class='background-lightpink'>";
                                    } else {
                                        echo "<tr>";
                                    }
                                        echo "<td class='padding-top-16 padding-bottom-16'>${price}</td>";
                                        $result = $mysqli->query("SELECT cutflowers${day} FROM route WHERE store_id = ${store_id}");
                                        $row = $result->fetch_row();
                                        $quantity = json_decode($row[0], true);
                                        $last = NULL;
                                        foreach ((array)$quantity as $key => $value) {
                                            if ($price == $key) {
                                                $last = $value;
                                                break 1;
                                            }
                                        }
                                        if ($price != "クリザ") {
                                            echo "<td class='padding-top-16 padding-bottom-16'><input type='number' name='quantity' min='0' max='255' value='${last}' class='cutflowers${day}_${store_id} _${price} width-100pe height-48'></td>";
                                        } else {
                                            echo "<td class='padding-top-16 padding-bottom-16'></td>";
                                        }
                                        $result = $mysqli->query("SELECT horticulture${day} FROM route WHERE store_id = ${store_id}");
                                        $row = $result->fetch_row();
                                        $quantity = json_decode($row[0], true);
                                        $last = NULL;
                                        foreach ((array)$quantity as $key => $value) {
                                            if ($price == $key) {
                                                $last = $value;
                                                break 1;
                                            }
                                        }
                                        if ($price != "クリザ" && $price != "ビター") {
                                            echo "<td class='padding-top-16 padding-bottom-16'><input type='number' name='quantity' min='0' max='255' value='${last}' class='horticulture${day}_${store_id} _${price} width-100pe height-48'></td>";
                                        } else {
                                            echo "<td class='padding-top-16 padding-bottom-16'></td>";
                                        }
                                        $result = $mysqli->query("SELECT cleyera${day} FROM route WHERE store_id = ${store_id}");
                                        $row = $result->fetch_row();
                                        $quantity = json_decode($row[0], true);
                                        $last = NULL;
                                        foreach ((array)$quantity as $key => $value) {
                                            if ($price == $key) {
                                                $last = $value;
                                                break 1;
                                            }
                                        }
                                        if ($price == 228 || $price == 358 || $price == 458 || $price == 550 || $price == 798) {
                                            echo "<td class='padding-top-16 padding-bottom-16'><input type='number' name='quantity' min='0' max='255' value='${last}' class='cleyera${day}_${store_id} _${price} width-100pe height-48'></td>";
                                        } else {
                                            echo "<td class='padding-top-16 padding-bottom-16'></td>";
                                        }
                                        $result = $mysqli->query("SELECT materials${day} FROM route WHERE store_id = ${store_id}");
                                        $row = $result->fetch_row();
                                        $quantity = json_decode($row[0], true);
                                        $last = NULL;
                                        foreach ((array)$quantity as $key => $value) {
                                            if ($price == $key) {
                                                $last = $value;
                                                break 1;
                                            }
                                        }
                                        echo "<td class='padding-top-16 padding-bottom-16'><input type='number' name='quantity' min='0' max='255' value='${last}' class='materials${day}_${store_id} _${price} width-100pe height-48'></td>";
                                    echo "</tr>";
                                }
                            echo "</table>
                        </div>";
                        $i++;
                    }
                echo "</div>
            </div>
            <div class='flex fixed space-between bottom-0 width-100-06 background-white'>
                <div class='width-19pe font-size-14'><a href='vertical.php?logout=true'>Logout</a></div>
                <div class='width-18pe'><strong>切花</strong></div>
                <div class='width-21pe'><strong>園芸</strong></div>
                <div class='width-16pe'><strong>榊</strong></div>
                <div><strong>資材</strong></div>
                <div class='none-767'>
                    <form action='vertical.php' method='post'>
                        <input type='submit' value='出力' name='upload'>
                    </form>
                </div>
            </div>";
        } else {
            echo "<div class='text-align-center'>
                <div class='padding-top-16 padding-bottom-16'>
                    <select>
                        <option hidden>選択してください</option>";
                        $course_name = [101, 102, 103, 104, 105, 201, 202, 203, 301, 302, 303, 304, 401, 402, 403, 404, 405, 501, 502, 503, 504, 505, 601, 602, 603, 604, 605];
                        foreach ($course_name as $value) {
                            echo "<option value='${value}'>${value}</option>";
                        }
                    echo "</select>
                </div>
                <div class='padding-top-16'>
                    <a href='inventory.php'>横スクロールに変更</a>
                </div>
            </div>";
        }
        ?>
        <script>
            (function($) {

                $(function() {
                    $('select').change(function() {
                        var val = $('option:selected').val();
                        $.post("inventory_data.php", {
                            "course_name":val
                        }, function(rs) {
                            location.href = 'http://54.88.35.31/vertical.php';
                        });
                    });
                });

                $(function() {
                    var index;
                    var nextIndex;
                    $('input').on("keydown", function(e) {
                        index = $('input').index(this);
                        var n = $("input").length;
                        if (e.which == 13) {
                            e.preventDefault();
                            nextIndex = $('input').index(this) + 1;
                            if (nextIndex < n) {
                                $('input').eq(nextIndex).focus();
                            } else {
                                $('input').eq(0).focus();
                            }
                        }
                    });
                 });

                 $(function() {
                    $("input[name='quantity']").change(function() {
                        $(this).css('color','red');
                        var val = $(this).val();
                        var store_price = $(this).attr('class');
                        var store_id = store_price.split(' ');
                        register(store_id[0]);
                    });
                });

                function register (num) {
                    var array = {};
                    $("." + num).each(function() {
                        var val = ($(this).val());
                        var store_price = $(this).attr('class');
                        var price = store_price.split(num);
                        var price = price[1].split(' _')
                        var price = price[1].split(' ')
                        array[price[0]] = val;
                    });
                    var store_id = num.split('_');
                    $.post("inventory_data.php", {
                        store_id: store_id[1], quantity: array , category : store_id[0]
                    }).done(function(data) {
                        $("." + num).each(function() {
                            $(this).css('color','');
                        });
                    });
                }

            })(jQuery);
        </script>
    </body>
</html>
