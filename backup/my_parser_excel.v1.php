<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>XLS Parser (D. Galinouski)</title>
    <style>
        * {
            margin: 0;
            padding: 0;
        }

        body {
            justify-content: center;
            align-items: center;
            padding: 20px;
            background: #f5ca8f;
        }
    </style>

</head>
<body>

<div>
    <h2>Парсинг данных excel документа</h2>
    <br>
    <h3>введите данные для выборки:</h3>
    <br>
    <form method="post" enctype="multipart/form-data">
        *.XLSX <input type="file" name="file"  />&nbsp;&nbsp;
<!--        <br><br>укажите поле: <input type="text" name="field" /><br>-->
        <br><input type="submit" value="Старт" /><br>
    </form>
</div>

<?php
//var_dump($_POST['field']); die;

function DBResult(&$_Result_id) { // Возвращение результата запроса
    $_Res = false;
    if ($_Result_id and ($_Result_id !== true)) {
        $_Res = array();
        while ($_Row = mysqli_fetch_assoc($_Result_id)) {
            $_Res[] = $_Row;
        }
    } elseif ($_Result_id === true) {
        $_Res = true;
    } else {
        return null;
    }
    return $_Res;
}

if ($_FILES['file']['tmp_name']) {
    $file_name = $_FILES['file']['tmp_name'];
    //$field = $_POST['field'];



    // Подключаем библиотеку
    require_once __DIR__ . "/PHPExcel/Classes/PHPExcel.php";
    // Подключаем модуль
    require_once __DIR__ . "/library/excel_mysql.php";

    // Определяем константу для включения режима отладки (режим отладки выключен)
    define("EXCEL_MYSQL_DEBUG", false);
    // Соединение с базой MySQL
    $connection = new mysqli("localhost", "root", "", "excel_mysql_base");

    // Выбираем кодировку UTF-8
    $connection->set_charset("utf8");

    // загрузка в базу данных оригинальной таблицы
    $excel_mysql_import = new Excel_mysql($connection, $file_name);
//    echo "<br>"."запись в базу данных 'original': ";
//    echo $excel_mysql_import->excel_to_mysql_by_index(
//        "original",
//        0,
//        array(
//            "id",
//            "name",
//            "price",
//            "store",
//        )
//    ) ? "OK\n" : "FAIL\n";
    echo "<hr>";


    $sql = "SELECT * FROM `original` WHERE price>1400 and price<2500";

    $res = mysqli_query($connection, $sql);
    $result_1 = DBResult($res);

    // сохранить копию оригинальной таблицы
    $objPHPExcel = new PHPExcel();
    $objPHPExcel = PHPExcel_IOFactory::load($file_name);
    $excel_writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $excel_writer -> save('./test.xls');

    $objPHPExcel2 = new PHPExcel();
    $objPHPExcel2 = PHPExcel_IOFactory::load('./test.xls');

    $objWorkSheet = new PHPExcel_Worksheet($objPHPExcel2, 'Данные выборки');
    $objPHPExcel2->addSheet($objWorkSheet, 0);
    //$get_field = $objPHPExcel2->getActiveSheet()->getCell($field)->getValue(); //получить заданное значение

    $objPHPExcel2->setActiveSheetIndexByName('Данные выборки'); // перейти на рабочий лист
    $objWorkSheet->getTabColor()->setRGB('FF0000');

    // Получить выборку 1

    $active_sheet = $objPHPExcel2->getActiveSheet();
    $active_sheet->getStyle('A1:D1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
    $active_sheet->getStyle('A1:D1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
    $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

    $active_sheet->setTitle('Данные выборки 1');
    $active_sheet->getColumnDimension('A')->setWidth(20);
    $active_sheet->getColumnDimension('B')->setWidth(20);
    $active_sheet->getColumnDimension('C')->setWidth(20);
    $active_sheet->getColumnDimension('D')->setWidth(20);

    //Вставка данных из выборки 1
    $start = 2;
    $i = 0;
    $active_sheet->setCellValueByColumnAndRow(0, 1, 'id');
    $active_sheet->setCellValueByColumnAndRow(1, 1, 'name');
    $active_sheet->setCellValueByColumnAndRow(2, 1, 'price');
    $active_sheet->setCellValueByColumnAndRow(3, 1, 'store');

    foreach ($result_1 as $row_1){
        $next = $start + $i;
        $active_sheet->setCellValueByColumnAndRow(0, $next, $row_1['id']);
        $active_sheet->setCellValueByColumnAndRow(1, $next, $row_1['name']);
        $active_sheet->setCellValueByColumnAndRow(2, $next, $row_1['price']);
        $active_sheet->setCellValueByColumnAndRow(3, $next, $row_1['store']);
        $i++;
    }

    // сохранить test
    $excel_writer = PHPExcel_IOFactory::createWriter($objPHPExcel2, 'Excel5');
    $excel_writer -> save('./test.xls');

    $objPHPExcel = new PHPExcel();
    $objPHPExcel = PHPExcel_IOFactory::load('./test.xls');
    $objWorksheet = $objPHPExcel->setActiveSheetIndexByName('Данные выборки 1'); // перейти на рабочий лист
    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'

    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5

    echo "<br><b> Результаты выборки 1 (1.xls) </b><br><br>";
    echo '<table style="width: 80%">' . "\n";
    for ($row = 1; $row <= $highestRow; ++$row) {
        echo '<tr>' . "\n";

        for ($col = 0; $col <= $highestColumnIndex; ++$col) {
            echo '<td style="width: 30%">' . $objWorksheet->getCellByColumnAndRow($col, $row)->getValue() . '</td>' . "\n";
        }

        echo '</tr>' . "\n";
    }
    echo '</table>' . "\n";

    //сохранить лист с выборкой 1
    $excel_writer = new PHPExcel_Writer_Excel5($objPHPExcel);
    $excel_writer->save("./1.xls");
    // сохранить в базу данных
    $excel_mysql_import = new Excel_mysql($connection, './1.xls');

    // Указываем названия столбцов в таблице MySQL
    echo "<br>"."запись в базу данных 'research_1': ";
    echo $excel_mysql_import->excel_to_mysql_by_index(
        "research_1",
        0,
        array(
            "id",
            "name",
            "price",
            "store",
            null,
        )
    ) ? "OK\n" : "FAIL\n";
    echo "<hr>";


    $_FILES['file']['tmp_name'] = "";

};

echo '</body></html>';







