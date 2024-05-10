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
    <h3>введите данные для парсинга:</h3>
    <br>
    <form method="post" enctype="multipart/form-data">
        *.XLSX <input type="file" name="file"  />&nbsp;&nbsp;
        <br><br>(Min) начальная цена: <input type="text" name="start_price" /><br>
        <br>(Max) максиммальная цена: <input type="text" name="high_price" /><br>
        <br>Наименование товара: <input type="text" name="name" /><br>
        <br>В наименовании товара присутствует текст: <input type="text" name="name_like" /><br>
        <br>Диапозон по атриклю от: <input type="text" name="id_start" /> до: <input type="text" name="id_finish" /><br>

        <br>Атрикль начинается с: <input type="text" name="id_n_start" /> оканчивается: <input type="text" name="id_n_finish" /><br>
        <br>колличество строк парсинга (все по умолчанию): <input type="text" name="limit" /><br>
        <br>
        Варианты парсинга согласно технического задания: <br>
        вариант 1: выбрать все строки у которых цена от Min и до Max<br>
        вариант 2: выбрать все строки у которых наименование товара Viatti Vettore Brina V-525 и цена больше Max включительно<br>
        вариант 3: выбрать все строки у которых артикул больше 100 000 но меньше 1000 000 включительно <br>
        вариант 4: выбрать все строки у которых артикул начинается с E4100 и заканчиваеться E4723 цена от 5500 до 7000 включительно,
        и наименование товара содержит: 104R<br>

        <br>выберите вариант выборки:  <select name="select">
                                            <option value="0"></option>
                                            <option value="1">вариант 1</option>
                                            <option value="2">вариант 2</option>
                                            <option value="3">вариант 3</option>
                                            <option value="4">вариант 4</option>
                                       </select>

        <br><br><input type="submit" value="Старт" /><br>
    </form>
</div>

<?php

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

if ($_FILES['file']['tmp_name'] && $_POST ) {
    $file_name = $_FILES['file']['name'];

    if ($_POST['start_price'])
        $start_price = $_POST['start_price'];
    else $start_price = 0;

    if ($_POST['high_price'])
        $high_price = $_POST['high_price'];
    else $high_price = 1000000;

    if ($_POST['id_start'])
        $id_start = $_POST['id_start'];
    else $id_start = 100000;

    if ($_POST['id_finish'])
        $id_finish = $_POST['id_finish'];
    else $id_finish = 1000000;

    if ($_POST['id_n_start'])
        $id_start = $_POST['id_n_start'];
    else $id_n_start = 'E4100';

    if ($_POST['id_n_finish'])
        $id_n_finish = $_POST['id_n_finish'];
    else $id_n_finish = 'E4723';


    if ($_POST['select'] && $_POST['select'] != 0 )
        $choice = $_POST['select'];
    else {
        echo "<br><b style='color: red'> не выбран вариант выборки!</b>";
        exit();
    }

    if ($_POST['limit'])
        $limit = $_POST['limit'];


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
    echo "<br>"."запись в базу данных файла: ".$file_name.". ";
    echo $excel_mysql_import->excel_to_mysql_by_index(
        "original",
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

    switch ($choice){
        case 1 :
            if($limit){
                $sql_1 = "SELECT * FROM `original` WHERE price> $start_price and price< $high_price LIMIT $limit ";
            }else $sql_1 = "SELECT * FROM `original` WHERE price> $start_price and price< $high_price ";

            $res_1 = mysqli_query($connection, $sql_1);
            $result = DBResult($res_1);  // результат выборки 1
            break;

        case 2 :
            if ($_POST['name'])
                $name = $_POST['name'];
            else {
                echo "<br><b style='color: red'> не выбрано наименование товара!</b>";
                exit();
            }
            if($limit){
                $sql_2 = "SELECT * FROM `original` WHERE name = '$name' and price >= $high_price LIMIT $limit " ;
            }else $sql_2 = "SELECT * FROM `original` WHERE name = '$name' and price >= $high_price " ;

            $res_2 = mysqli_query($connection, $sql_2);
            $result = DBResult($res_2);  // результат выборки 2
            break;

        case 3 :
            if($limit) {
                $sql_3 = "SELECT * FROM `original` WHERE id > $id_start AND id <= $id_finish LIMIT $limit ";
            }else $sql_3 = "SELECT * FROM `original` WHERE id > $id_start AND id <= $id_finish ";

            $res_3 = mysqli_query($connection, $sql_3);
            $result = DBResult($res_3);  // результат выборки 3
            break; //SELECT * FROM `original` WHERE id BETWEEN 'E4100' AND 'E4723' AND price BETWEEN 5500 and 7000 AND name LIKE '%104R%'

        case 4 :

            if ($_POST['name_like'])
                $name_like = $_POST['name_like'];
            else {
                echo "<br><b style='color: red'> не выбран предположительный текст в наименовании товара!</b>";
                exit();
            }
            if($limit) {
                $sql_4 = "SELECT * FROM `original` WHERE id BETWEEN '$id_n_start' AND '$id_n_finish' AND price BETWEEN $start_price and $high_price AND name LIKE '%$name_like%' LIMIT $limit ";
            }else $sql_4 = "SELECT * FROM `original` WHERE id BETWEEN 'E4100' AND 'E4723' AND price BETWEEN $start_price and $high_price AND name LIKE '%$name_like%'";

            $res_4 = mysqli_query($connection, $sql_4);
            $result = DBResult($res_4);  // результат выборки 4
            break;
    }

    if($result == NULL){
        echo "<br><b style='color: red'>Проверьте введённые данные!</b>";
        exit();
    }


    // старт работы
    $objPHPExcel = new PHPExcel();
    $objPHPExcel = PHPExcel_IOFactory::load($file_name);

    $objWorkSheet = new PHPExcel_Worksheet($objPHPExcel, 'Данные выборки');
    $objPHPExcel->addSheet($objWorkSheet, 0);

    $objPHPExcel->setActiveSheetIndexByName('Данные выборки'); // перейти на рабочий лист
    $objWorkSheet->getTabColor()->setRGB('FF0000');
    $active_sheet = $objPHPExcel->getActiveSheet();
    $active_sheet->getStyle('A1:D1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);
    $active_sheet->getStyle('A1:D1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

    $active_sheet->getColumnDimension('A')->setWidth(20);
    $active_sheet->getColumnDimension('B')->setWidth(40);
    $active_sheet->getColumnDimension('C')->setWidth(20);
    $active_sheet->getColumnDimension('D')->setWidth(20);

    //Вставка данных из выборки 1
    $start = 2;
    $i = 0;
    $active_sheet->setCellValueByColumnAndRow(0, 1, 'id');
    $active_sheet->setCellValueByColumnAndRow(1, 1, 'name');
    $active_sheet->setCellValueByColumnAndRow(2, 1, 'price');
    $active_sheet->setCellValueByColumnAndRow(3, 1, 'store');

    foreach ($result as $row){
        $next = $start + $i;
        $active_sheet->setCellValueByColumnAndRow(0, $next, $row['id']);
        $active_sheet->setCellValueByColumnAndRow(1, $next, $row['name']);
        $active_sheet->setCellValueByColumnAndRow(2, $next, $row['price']);
        $active_sheet->setCellValueByColumnAndRow(3, $next, $row['store']);
        $i++;
    }

//сохранить лист с выборкой в файл research.xls
    $excel_writer = new PHPExcel_Writer_Excel5($objPHPExcel);
    $excel_writer->save("./research.xls");
    $objPHPExcel = new PHPExcel();
    $objPHPExcel = PHPExcel_IOFactory::load('./research.xls');

    $objWorksheet = $objPHPExcel->setActiveSheetIndex(); // перейти на рабочий лист
    $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
    $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'

    $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn); // e.g. 5
// вывод на экран
    echo "<br><b> Результаты выборки (research.xls) </b><br><br>";
    echo '<table style="width: 80%">' . "\n";
    for ($row = 1; $row <= $highestRow; ++$row) {
        echo '<tr>' . "\n";

        for ($col = 0; $col <= $highestColumnIndex; ++$col) {
            echo '<td style="width: 30%">' . $objWorksheet->getCellByColumnAndRow($col, $row)->getValue() . '</td>' . "\n";
        }

        echo '</tr>' . "\n";
    }
    echo '</table>' . "\n";

    // сохранить в базу данных
    $excel_mysql_import = new Excel_mysql($connection, 'research.xls');
    // Указываем названия столбцов в таблице MySQL
    echo "<br>"."запись в базу данных 'research': ";
    echo $excel_mysql_import->excel_to_mysql_by_index(
        "research",
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



