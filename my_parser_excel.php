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

        <br>или начинается с: <input type="text" name="id_n_start" /> и оканчивается на: <input type="text" name="id_n_finish" /><br>
        <br>колличество строк парсинга (все по умолчанию): <input type="text" name="limit" /><br>

        <br><br><input type="submit" value="Старт" /><br>
    </form>
</div>

<?php

// Подключаем библиотеку

require_once __DIR__ . "/PHPExcel/Classes/PHPExcel.php";
// Подключаем модуль
require_once __DIR__ . "/library/excel_mysql.php";

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

    if ($_POST['high_price'])
        $high_price = $_POST['high_price'];

    if ($_POST['id_start'])
        $id_start = $_POST['id_start'];

    if ($_POST['id_finish'])
        $id_finish = $_POST['id_finish'];

    if ($_POST['id_n_start'])
        $id_n_start = $_POST['id_n_start'];

    if ($_POST['id_n_finish'])
        $id_n_finish = $_POST['id_n_finish'];

    if ($_POST['name'])
        $name = $_POST['name'];

    if ($_POST['name_like'])
        $name_like = $_POST['name_like'];

    if ($_POST['limit'])
        $limit = $_POST['limit'];

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


            $sql = "SELECT * FROM `original` WHERE price>0 ";

            if (isset($name) && !isset($name_like))
                $sql .= " AND name = '$name'";
            if (isset($name_like) && !isset($name))
                $sql .= " AND name LIKE '%$name_like%' ";
            if (isset($name_like) && isset($name))
                $sql .= " AND name = '$name' OR name LIKE '%$name_like%' ";
            if (isset($id_n_start) && !isset($id_n_finish))
                $sql .= " AND id LIKE '%$id_n_start%' ";
            if (isset($id_n_finish) && !isset($id_n_start))
                $sql .= " AND id LIKE '%$id_n_finish%' ";
            if (isset($id_n_start) && isset($id_n_finish))
                $sql .= " AND id BETWEEN '$id_n_start' AND '$id_n_finish' ";
            if (isset($id_start) && !isset($id_finish))
                $sql .= " AND id > '$id_start' ";
            if (isset($id_finish) && !isset($id_start))
                $sql .= " AND id <= '$id_finish' ";
            if (isset($id_finish) && isset($id_start))
                $sql .= " AND id BETWEEN '$id_start' AND '$id_finish' ";
            if (isset($start_price))
                $sql .= " AND price > $start_price ";
            if (isset($high_price))
                $sql .= " AND price <= $high_price ";
            if (isset($limit))
                $sql .= " LIMIT $limit";

    $res = mysqli_query($connection, $sql);
    if(!$res){
        echo "<br><b style='color: red'>Проверьте введённые данные!</b>";
        exit();
    }

    $result = DBResult($res);  // результат выборки

    if($result == NULL){
        echo "<br><b style='color: red'>Проверьте введённые данные!</b>";
        exit();
    }


    // старт работы
    $objPHPExcel = new PHPExcel();

    $active_sheet = $objPHPExcel->getActiveSheet();
    $active_sheet->setTitle('Данные выборки');
    $active_sheet->getTabColor()->setRGB('FF0000');
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
        if ($row['id'] == NULL) continue;
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

    $objWorksheet = $objPHPExcel->setActiveSheetIndex(0); // перейти на рабочий лист
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



