<?php

require_once __DIR__ . "./vendor/autoload.php";
global $pdo;

require_once __DIR__ . "./configs/database_config.php";
require_once __DIR__ . "./library/database.php";
require_once __DIR__ . "./library/functions.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;


if ($_FILES['file']['tmp_name'] && $_POST ) {
    $fileName = $_FILES['file']['name'];

    if ($_POST['startPrice'])
        $startPrice = $_POST['startPrice'];
    else $startPrice = '';

    if ($_POST['highPrice'])
        $highPrice = $_POST['highPrice'];
    else $highPrice = '';

    if ($_POST['idStart'])
        $idStart = $_POST['idStart'];
    else $idStart = '';

    if ($_POST['idFinish'])
        $idFinish = $_POST['idFinish'];
    else $idFinish = '';

    if ($_POST['idTitleBegin'])
        $idTitleBegin = $_POST['idTitleBegin'];
    else $idTitleBegin = '';

    if ($_POST['idTitleAnd'])
        $idTitleAnd = $_POST['idTitleAnd'];
    else $idTitleAnd = '';

    if ($_POST['name'])
        $name = $_POST['name'];
    else $name = '';

    if ($_POST['nameLike'])
        $nameLike = $_POST['nameLike'];
    else $nameLike = '';

    if ($_POST['limit'])
        $limit = $_POST['limit'];
    else $limit = '';

    $spreadsheet = new Spreadsheet();
    $spreadsheet = readingXls($fileName);
    $spreadsheet->setActiveSheetIndex(0);
    $sheet = $spreadsheet->getActiveSheet();

    // замеряем время работы скрипта
    $startTime = microtime(true);

    // запускаем экспорт данных
    excel2db($spreadsheet, $pdo, 'original');
    $elapsedTime = round(microtime(true) - $startTime, 4);
    echo "<br><b>Загрузка в базу данных 'original': $elapsedTime с.</b><br>";

    $sql = "SELECT * FROM `original` WHERE price>0 ";

    if ($name!='' && $nameLike=='')
        $sql .= " AND name = '$name' ";
    if ($nameLike!='' && $name=='')
        $sql .= " AND name LIKE '% $nameLike %' ";
    if ($nameLike!='' && $name!='')
        $sql .= " AND name = '$name' OR name LIKE '% $nameLike %' ";
    if ($idTitleBegin!='' && $idTitleAnd=='')
        $sql .= " AND id LIKE '% $idTitleBegin %' ";
    if ($idTitleAnd!='' && $idTitleBegin=='')
        $sql .= " AND id LIKE '% $idTitleAnd %' ";
    if ($idTitleBegin=='' && $idTitleAnd!='')
        $sql .= " AND id BETWEEN ' $idTitleBegin ' AND ' $idTitleAnd ' ";
    if ($idStart!='' && $idFinish=='')
        $sql .= " AND id > ' $idStart ' ";
    if ($idFinish!='' && $idStart=='')
        $sql .= " AND id <= ' $idFinish ' ";
    if ($idFinish!='' && $idStart!='')
        $sql .= " AND id BETWEEN ' $idStart ' AND ' $idFinish ' ";
    if ($startPrice!='')
        $sql .= " AND price > $startPrice ";
    if ($highPrice!='')
        $sql .= " AND price <= $highPrice ";
    if ($limit!='')
        $sql .= " LIMIT $limit";

    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $result = $stmt->fetchAll(PDO::FETCH_ASSOC);

    if(!$result){
        echo "<br><br><span style='color: red'>Пожалуйста, проверьте введённые данные.</span>";
    }

    // старт работы c новым xls файлом в котором будет храниться выборка
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

    //Вставка данных выборки в xls файл
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
    savingXls($spreadsheet, 'research.xls');

    // чтение файла 'research.xls'
    $spreadsheet = readingXls('research.xls');
    $spreadsheet->setActiveSheetIndex(0);
    $sheet = $spreadsheet->getActiveSheet();

    // вывод данных в шаблон results.php
    $html = showResults($sheet);
    include_once dirname(__FILE__).'./templates/results.php';

    // сохранить в базу данных результат выборки
    // замеряем время работы скрипта
    $startTime = microtime(true);
    excel2db($spreadsheet, $pdo, 'research');
    $elapsedTime = round(microtime(true) - $startTime, 4);
    echo "<br><b>Загрузка в базу данных 'research': $elapsedTime с.</b><br>";

    $dsn = NULL;
    $_FILES['file']['tmp_name'] = "";

};


