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
require 'vendor/autoload.php';
// Подключаем модуль
require_once __DIR__ . "/library/excel2db.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\IOFactory;


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


    // Соединение с базой MySQL

    $host = 'localhost';
    $db   = 'excel_mysql_base';
    $user = 'root';
    $pass = '';
    $charset = 'utf8';

    $dsn = "mysql:host=$host;dbname=$db;charset=$charset";
    $opts= [
        \PDO::ATTR_ERRMODE            => \PDO::ERRMODE_EXCEPTION,
        \PDO::ATTR_DEFAULT_FETCH_MODE => \PDO::FETCH_ASSOC,
        \PDO::ATTR_EMULATE_PREPARES   => false,
    ];

    // подключение к базе
    $pdo = new PDO($dsn, $user, $pass, $opts);

    // класс, который читает xls файл
    $spreadsheet = new Spreadsheet();
    $reader = new Xls($spreadsheet);
    // получаем Excel-книгу
    $reader->load($file_name);

    // замеряем время работы скрипта
    $startTime = microtime(true);
    // запускаем экспорт данных
    $table = 'original';
    excel2db($spreadsheet, $pdo, $table, false);
    $elapsedTime = round(microtime(true) - $startTime, 4);
    echo "<br><b>Загрузка в базу данных: $elapsedTime с.</b><br>";


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

    $connection = new mysqli("localhost", "root", "", "excel_mysql_base");
    // Выбираем кодировку UTF-8
    $connection->set_charset($charset);

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
    $spreadsheet = new Spreadsheet();

    $active_sheet = $spreadsheet->getActiveSheet();
    $active_sheet->setTitle('Данные выборки');
    $active_sheet->getTabColor()->setRGB('FF0000');
    $active_sheet->getStyle('A1:D1')->applyFromArray([
        'font' => [
            'name' => 'Arial',
            'bold' => true,
            'italic' => false,
            'underline' => Font::UNDERLINE_DOUBLE,
            'strikethrough' => false,
            'color' => [
                'rgb' => 'red'
            ]
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => [
                    'rgb' => 'black'
                ]
            ],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
            'wrapText' => true,
        ]
    ]);


    $active_sheet->getColumnDimension('A')->setWidth(20);
    $active_sheet->getColumnDimension('B')->setWidth(40);
    $active_sheet->getColumnDimension('C')->setWidth(20);
    $active_sheet->getColumnDimension('D')->setWidth(20);

    //Вставка данных из выборки 1
    $start = 2;
    $i = 0;
    $active_sheet->setCellValue('A1', 'id');
    $active_sheet->setCellValue('B1', 'name');
    $active_sheet->setCellValue('C1', 'price');
    $active_sheet->setCellValue('D1', 'store');

    foreach ($result as $row){
        $next = $start + $i;
        if ($row['id'] == NULL) continue;
        $active_sheet->setCellValue('A'.$next, $row['id']);
        $active_sheet->setCellValue('B'.$next, $row['name']);
        $active_sheet->setCellValue('C'.$next, $row['price']);
        $active_sheet->setCellValue('D'.$next, $row['store']);
        $i++;
    }

    //сохранить лист с выборкой в файл research.xls
    // Выбросим исключение в случае, если не удастся сохранить файл
    try {
        $writer = new Xls($spreadsheet);
        $writer->save('./research.xls');

    } catch (PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
        echo $e->getMessage();
    }

    $reader = IOFactory::createReader('Xls');
    $spreadsheet = $reader->load('./research.xls');
    // Количество листов
    $spreadsheet->setActiveSheetIndex(0); // получить данные из указанного листа
    $sheet = $spreadsheet->getActiveSheet();

    // формирование html-кода с данными
    $html = '<br><table style="width: 70%;">';
    foreach ($sheet->getRowIterator() as $row) {
        $html .= '<tr>';
        $cellIterator = $row->getCellIterator();
        foreach ($cellIterator as $cell) {

            // значение текущей ячейки
            $value = $cell->getCalculatedValue();

            $html .= '<td>'.$value.'</td>';
        }
        $html .= '<tr>';
    }
    $html .= '</table>';

    // вывод данных
    echo $html;

    // сохранить в базу данных результат выборки
    // замеряем время работы скрипта
    $startTime = microtime(true);
    // запускаем экспорт данных
    // подключение к базе

    $table = 'research';
    excel2db($spreadsheet, $pdo, $table, false);
    $elapsedTime = round(microtime(true) - $startTime, 4);
    echo "<br><b>Загрузка в базу данных: $elapsedTime с.</b><br>";


    $_FILES['file']['tmp_name'] = "";

};

echo '</body></html>';



